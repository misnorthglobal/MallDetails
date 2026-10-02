"""UAE Mall Dashboard
Install: python -m pip install streamlit pandas openpyxl
Run: python -m streamlit run mall_dashboard.py
Put 2gis_all_malls_company_details.xlsx beside this script, or upload it.
Requires Python 3.10+ and Streamlit 1.35+.
"""
from pathlib import Path
import re
import pandas as pd


def clean(value):
    return '' if pd.isna(value) else re.sub(r'\s+', ' ', str(value).replace('\u200b', ' ')).strip()


def prepare(source):
    sheets = pd.read_excel(source, sheet_name=['Malls', 'Companies'])
    m, c = sheets['Malls'], sheets['Companies']
    required = {'Malls': ['Mall Number', 'Name', 'Approximate Size', 'Emirates', 'Remarks'],
                'Companies': ['Mall Number', 'Company Name', 'Category', 'Company URL']}
    for name, df in sheets.items():
        df.columns = [clean(x) for x in df.columns]
        missing = set(required[name]) - set(df.columns)
        if missing:
            raise ValueError(f'{name}: missing columns {sorted(missing)}')
        for col in df.select_dtypes(include=['object', 'string']):
            df[col] = df[col].map(clean)
        df['Mall Number'] = pd.to_numeric(df['Mall Number'], errors='coerce').astype('Int64')
    if m['Mall Number'].isna().any() or m['Mall Number'].duplicated().any():
        raise ValueError('Malls must have unique, nonempty Mall Number values.')
    for col in ['Property Owner', 'Confidence', 'Location', 'Size Range', 'Link']:
        if col not in m:
            m[col] = ''
    for col in ['Mall Name', 'Emirates', 'Floors', 'Branches Text', 'Remarks']:
        if col not in c:
            c[col] = ''
    for col in ['Companies', 'Floors', 'Total Parking']:
        m[col] = pd.to_numeric(m.get(col, pd.Series(index=m.index, dtype=float)), errors='coerce')
    c['Branch Counts'] = pd.to_numeric(c.get('Branch Counts', pd.Series(index=c.index, dtype=float)), errors='coerce')
    c['Category'] = c['Category'].replace('', 'Uncategorized')
    m['Size (sq ft)'] = pd.to_numeric(m['Approximate Size'].str.replace(',', '', regex=False)
                                   .str.extract(r'(\d+(?:\.\d+)?)')[0], errors='coerce')
    m['Status'] = m['Remarks'].map(lambda x: 'Coming soon' if re.search(r'com(?:ing|ign)\s*soon', x, re.I)
                                  else 'Status not specified')
    counts = c.groupby('Mall Number').size()
    m['Extracted records'] = m['Mall Number'].map(counts).fillna(0).astype(int)
    m['Company count difference'] = m['Extracted records'] - m['Companies']
    return m, c


def rule_mask(df, column, operator, value):
    s = df[column].fillna('').astype(str).map(clean).str.casefold()
    q = clean(value).casefold()
    if operator == 'Is empty':
        return s.eq('')
    if operator == 'Is not empty':
        return s.ne('')
    if operator == 'Equals':
        return s.eq(q)
    if operator == 'Starts with':
        return s.str.startswith(q)
    hit = s.str.contains(q, regex=False)
    return ~hit if operator == 'Does not contain' else hit


def apply_rules(df, rules, mode):
    masks = [rule_mask(df, *r) for r in rules if r[1] in ['Is empty', 'Is not empty'] or clean(r[2])]
    if not masks:
        return df
    mask = masks[0].copy()
    for next_mask in masks[1:]:
        mask = mask & next_mask if mode == 'All rules (AND)' else mask | next_mask
    return df.loc[mask]


def main():
    import streamlit as st
    st.set_page_config(page_title='UAE Mall Dashboard', page_icon='🏙️', layout='wide')
    st.title('UAE Mall Dashboard')
    st.caption('Recorded Excel estimates and statuses. Company records are listings, not unique businesses.')
    default = Path(__file__).parent / '2gis_all_malls_company_details.xlsx'
    if not default.exists():
        default = Path(__file__).parent / 'data' / default.name
    with st.sidebar:
        uploaded = st.file_uploader('Excel workbook', type=['xlsx'])
        if st.button('Reset filters'):
            for key in list(st.session_state):
                if key.startswith('f_'):
                    del st.session_state[key]
            st.rerun()
    @st.cache_data(show_spinner=False)
    def cached_load(content):
        from io import BytesIO
        return prepare(BytesIO(content))
    try:
        content = uploaded.getvalue() if uploaded else default.read_bytes()
        malls, companies = cached_load(content)
    except Exception as exc:
        st.error(f'Could not load workbook: {exc}. Upload your Excel using the sidebar.')
        st.stop()

    def select(df, col, label, key):
        opts = sorted(df[col].dropna().astype(str).unique())
        selected = st.multiselect(label, [x for x in opts if x], key=key)
        return df[df[col].isin(selected)] if selected else df

    def numeric(df, col, label, key):
        with st.expander(label):
            enabled = st.checkbox('Apply range', key=key+'_enabled')
            valid = df[col].dropna()
            maximum = float(valid.max()) if len(valid) else 0.0
            lower = st.number_input('Minimum', min_value=0.0, value=0.0, key=key+'_min')
            upper = st.number_input('Maximum', min_value=0.0, value=max(maximum, 1.0), key=key+'_max')
            unknown = st.checkbox('Include missing values', value=True, key=key+'_unknown')
            if enabled:
                if lower > upper:
                    st.warning('Minimum exceeds maximum.')
                df = df[df[col].between(lower, upper) | (df[col].isna() & unknown)]
        return df

    def searches(df, cols, prefix):
        with st.expander('Column-specific text search'):
            mode = st.radio('Combine search rules', ['All rules (AND)', 'Any rule (OR)'], key=prefix+'_mode')
            n = st.number_input('Number of rules', min_value=0, max_value=8, value=0, key=prefix+'_n')
            rules = []
            for i in range(int(n)):
                a, b, d = st.columns([2, 2, 3])
                col = a.selectbox(f'Column {i+1}', cols, key=f'{prefix}_{i}_col')
                op = b.selectbox(f'Condition {i+1}', ['Contains', 'Equals', 'Starts with', 'Does not contain',
                                                       'Is empty', 'Is not empty'], key=f'{prefix}_{i}_op')
                val = d.text_input(f'Text {i+1}', disabled=op in ['Is empty', 'Is not empty'], key=f'{prefix}_{i}_val')
                rules.append((col, op, val))
        return apply_rules(df, rules, mode)

    def table(df, key, height=420):
        config = {col: st.column_config.LinkColumn(col) for col in ['Link', 'Company URL'] if col in df}
        st.dataframe(df, hide_index=True, use_container_width=True, height=height, column_config=config)
        st.download_button('Download these results (CSV)', df.to_csv(index=False).encode('utf-8-sig'),
                           f'{key}.csv', 'text/csv', key='download_'+key)

    def bar_counts(df, col, title, limit=15):
        st.subheader(title)
        counts = df[col].replace('', pd.NA).dropna().value_counts().head(limit)
        if counts.empty:
            st.info('No records for this chart.')
        else:
            st.bar_chart(counts.rename('Records'))

    with st.sidebar:
        st.header('Mall filters')
        fm = malls.copy()
        for col in ['Emirates', 'Status', 'Property Owner', 'Confidence']:
            fm = select(fm, col, col, 'f_m_'+col)
        for col in ['Size (sq ft)', 'Floors', 'Total Parking', 'Extracted records']:
            fm = numeric(fm, col, col, 'f_mrange_'+col)
        coverage = st.selectbox('Company data', ['All', 'Has company records', 'No company records'], key='f_coverage')
        if coverage != 'All':
            fm = fm[fm['Extracted records'].gt(0) if coverage == 'Has company records' else fm['Extracted records'].eq(0)]
        cats = st.multiselect('Malls containing categories', sorted(companies.Category.unique()), key='f_cats')
        cat_mode = st.radio('Category match', ['Any selected category', 'All selected categories'], key='f_catmode')
        if cats:
            matches = companies[companies.Category.isin(cats)]
            ids = matches['Mall Number'].unique() if cat_mode.startswith('Any') else matches.groupby('Mall Number').Category.nunique().loc[lambda s:s.eq(len(cats))].index
            fm = fm[fm['Mall Number'].isin(ids)]
    fm = searches(fm, ['Name', 'Location', 'Emirates', 'Property Owner', 'Confidence', 'Remarks'], 'f_msearch')
    fm = fm.sort_values('Size (sq ft)', ascending=False, na_position='last')
    pool = companies[companies['Mall Number'].isin(fm['Mall Number'])].copy()
    st.caption(f'Active selection: {len(fm):,} / {len(malls):,} malls; {len(pool):,} / {len(companies):,} company records. Sidebar filters and mall text rules apply to every tab.')
    st.info('Blank remarks mean status not specified. Coming-soon malls without extracted listings have unavailable tenant data. Size definitions and parking values require validation before density or total-area comparisons.')
    tabs = st.tabs(['Overview', 'Mall rankings', 'Mall explorer', 'Company search', 'Compare malls', 'Data quality'])
    mall_cols = ['Mall Number', 'Name', 'Emirates', 'Status', 'Approximate Size', 'Size (sq ft)', 'Size Range',
                 'Property Owner', 'Confidence', 'Total Parking', 'Floors', 'Companies', 'Extracted records', 'Remarks', 'Location', 'Link']
    with tabs[0]:
        first, second, third = st.columns(3)
        first.metric('Total malls', f'{len(fm):,}')
        second.metric('Coming-soon malls', f'{fm.Status.eq("Coming soon").sum():,}')

        third.metric('Company records', f'{len(pool):,}')

        st.subheader('Malls by emirate')
        emirates = (fm.assign(Emirates=fm['Emirates'].replace('', 'Not specified'))
                    .groupby('Emirates').size().rename('Mall count').reset_index()
                    .sort_values(['Mall count', 'Emirates'], ascending=[False, True]))
        emirates = pd.concat([emirates, pd.DataFrame([{'Emirates': 'Total', 'Mall count': len(fm)}])], ignore_index=True)
        table(emirates, 'overview_emirates', 320)

        st.subheader('Top 10 property owners by mall count')
        owners = fm.assign(**{'Property Owner': fm['Property Owner'].replace('', 'Not publicly disclosed'),
                             'Coming-soon count': fm['Status'].eq('Coming soon').astype(int)})
        owners = (owners.groupby('Property Owner').agg(
            **{'Mall count': ('Mall Number', 'size'), 'Coming-soon count': ('Coming-soon count', 'sum')})
            .reset_index().sort_values(['Mall count', 'Property Owner'], ascending=[False, True]))
        table(owners.head(10), 'overview_owners')

        st.subheader('Top 10 malls by company count')
        ranking = fm.sort_values(['Extracted records', 'Name'], ascending=[False, True]).copy()
        ranking['Company records'] = ranking['Extracted records'].astype(object)
        ranking.loc[ranking['Status'].eq('Coming soon') & ranking['Extracted records'].eq(0), 'Company records'] = 'Not available'
        ranking = ranking.rename(columns={'Name': 'Mall name', 'Emirates': 'Emirate', 'Property Owner': 'Property owner'})
        table(ranking.head(10)[['Mall name', 'Emirate', 'Property owner', 'Status', 'Company records']], 'overview_company_counts', 550)
        st.caption('Company count means extracted listing records. Coming-soon malls without records show Not available.')
        st.subheader('Top 10 malls by approximate size')
        table(fm.nlargest(10, 'Size (sq ft)')[['Name', 'Emirates', 'Status', 'Approximate Size', 'Size Range']], 'overview_size')
    with tabs[1]:
        sort_col = st.selectbox('Sort by', ['Extracted records', 'Size (sq ft)', 'Total Parking', 'Floors', 'Name'], key='f_sort')
        ascending = st.checkbox('Ascending order', key='f_ascending')
        table(fm.sort_values(sort_col, ascending=ascending, na_position='last')[mall_cols], 'mall_rankings')
    with tabs[2]:
        st.subheader('Property owners')
        owner_data = fm.assign(**{'Property Owner': fm['Property Owner'].replace('', 'Not publicly disclosed'),
                                 'Coming soon': fm['Status'].eq('Coming soon').astype(int)})
        owner_summary = owner_data.groupby('Property Owner').agg(
            **{'Mall count': ('Mall Number', 'size'), 'Company records': ('Extracted records', 'sum'),
               'Coming-soon malls': ('Coming soon', 'sum')}).reset_index().sort_values('Mall count', ascending=False)
        table(owner_summary, 'explorer_owners')
        st.caption('Company records are listings inside the owner’s malls, not businesses legally owned by the owner.')
        st.subheader('Malls and company records by emirate')
        emirate_summary = fm.groupby('Emirates').agg(**{'Mall count': ('Mall Number', 'size'),
            'Company records': ('Extracted records', 'sum')}).reset_index().sort_values('Mall count', ascending=False)
        table(emirate_summary, 'explorer_emirates')
        chosen_owner = st.selectbox('Explore property owner', ['All owners'] + owner_summary['Property Owner'].tolist(), key='f_exploreowner')
        explore = fm if chosen_owner == 'All owners' else fm[fm['Property Owner'].replace('', 'Not publicly disclosed').eq(chosen_owner)]
        table(explore[mall_cols], 'owner_malls')
        if explore.empty:
            st.info('No matching malls.')
        else:
            ids = explore['Mall Number'].tolist()
            labels = explore.set_index('Mall Number').apply(lambda r: f'{r["Name"]} — {r["Emirates"]}', axis=1).to_dict()
            ident = st.selectbox('Choose mall', ids, format_func=lambda x: labels[x], key='f_explorer')
            row = explore[explore['Mall Number'].eq(ident)].iloc[0]
            st.subheader(row['Name'])
            st.write(row[['Status', 'Location', 'Property Owner', 'Approximate Size', 'Size Range', 'Confidence', 'Total Parking', 'Floors', 'Remarks']])
            if str(row['Link']).startswith(('https://', 'http://')):
                st.link_button('Open in 2GIS', row['Link'])
            tenants = pool[pool['Mall Number'].eq(ident)]
            if tenants.empty:
                st.info('No company records available in the workbook for this mall.')
            else:
                tenants = select(tenants, 'Category', 'Directory categories', 'f_dircats')
                st.metric('Company records', f'{len(tenants):,}')
                bar_counts(tenants, 'Category', 'Selected mall categories', limit=len(tenants))
                bar_counts(tenants, 'Floors', 'Selected mall company floors', limit=len(tenants))
                table(tenants, 'mall_directory')
    with tabs[3]:
        st.caption('These additional filters affect this company view only.')
        cp = select(pool, 'Mall Name', 'Malls', 'f_companymalls')
        cp = select(cp, 'Category', 'Categories', 'f_companycats')
        cp = select(cp, 'Floors', 'Company floors', 'f_companyfloors')
        cp = numeric(cp, 'Branch Counts', 'Reported branch count', 'f_branchrange')
        link_mode = st.selectbox('Company link', ['All', 'Available', 'Missing'], key='f_links')
        if link_mode != 'All':
            cp = cp[cp['Company URL'].ne('') if link_mode == 'Available' else cp['Company URL'].eq('')]
        cp = searches(cp, ['Company Name', 'Category', 'Mall Name', 'Emirates', 'Floors', 'Branches Text', 'Remarks'], 'f_csearch')
        a, b, d = st.columns(3)
        a.metric('Matching company records', f'{len(cp):,}')
        b.metric('Distinct company-name labels', f'{cp["Company Name"].nunique():,}')
        d.metric('Malls containing matches', f'{cp["Mall Number"].nunique():,}')
        st.caption(f'Company links: {cp["Company URL"].ne("").sum():,} available; {cp["Company URL"].eq("").sum():,} missing.')
        for group in ['Category', 'Mall Name', 'Emirates']:
            summary = cp.groupby(group).size().rename('Company records').reset_index().sort_values('Company records', ascending=False)
            st.subheader('Company records by ' + group.lower())
            table(summary, 'company_group_' + group)
        known = cp['Branch Counts'].dropna()
        st.caption(f'Known branch counts: {len(known):,}; multi-branch records: {known.gt(1).sum():,}. Branch totals are not summed across listings.')
        table(cp, 'company_results', 550)
        bar_counts(cp, 'Floors', 'Company records by floor')
        bar_counts(cp, 'Company Name', 'Most frequent company-name labels')
    with tabs[4]:
        ids = fm['Mall Number'].tolist()
        labels = fm.set_index('Mall Number').apply(lambda r: f'{r["Name"]} — {r["Emirates"]} (#{r.name})', axis=1).to_dict()
        selected = st.multiselect('Select up to 8 malls', ids, format_func=lambda x: labels[x], max_selections=8, key='f_compare')
        comparison = fm[fm['Mall Number'].isin(selected)]
        if len(comparison) < 2:
            st.info('Select at least two malls to compare.')
        if len(comparison):
            cat_pool = pool[pool.Category.ne('Uncategorized')]
            comparison = comparison.copy()
            comparison['Distinct categories'] = comparison['Mall Number'].map(cat_pool.groupby('Mall Number').Category.nunique()).fillna(0).astype(int)
            table(comparison[mall_cols + ['Distinct categories']], 'mall_comparison')
            mix = pool[pool['Mall Number'].isin(selected)].groupby(['Mall Number', 'Category']).size().rename('Company records').reset_index()
            mix['Mall'] = mix['Mall Number'].map(labels)
            st.subheader('Category mix')
            table(mix[['Mall', 'Category', 'Company records']], 'comparison_categories')
            for col in ['Size (sq ft)', 'Total Parking', 'Floors', 'Extracted records', 'Distinct categories']:
                st.subheader(col)
                chart = comparison.copy()
                chart['Mall'] = chart['Mall Number'].map(labels)
                st.bar_chart(chart.set_index('Mall')[[col]])
    with tabs[5]:
        st.metric('Malls without company records', f'{fm["Extracted records"].eq(0).sum():,}')
        rows = []
        for label, df, cols in [('Malls', fm, ['Property Owner', 'Size (sq ft)', 'Floors', 'Total Parking']),
                                ('Companies', pool, ['Category', 'Floors', 'Branch Counts', 'Company URL'])]:
            for col in cols:
                available = df[col].notna() & df[col].astype(str).str.strip().ne('')
                if col == 'Property Owner':
                    available &= ~df[col].str.contains('not publicly|undisclosed', case=False, na=False)
                if col == 'Category':
                    available &= df[col].ne('Uncategorized')
                count = int(available.sum())
                rows.append({'Sheet': label, 'Field': col, 'Available': count, 'Missing / undisclosed': len(df)-count,
                             'Coverage %': round(count/len(df)*100, 1) if len(df) else None})
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
        valid_links = pool.loc[pool['Company URL'].ne(''), 'Company URL']
        st.write(f'Duplicate nonempty company URLs: {valid_links.duplicated().sum():,}')
        st.write(f'Company records without a matching mall in the entire workbook: {(~companies["Mall Number"].isin(malls["Mall Number"])).sum():,}')
        st.caption('Source Companies count is compared with extracted row count. A difference indicates a coverage discrepancy, not vacant units.')
        table(fm[['Mall Number', 'Name', 'Status', 'Companies', 'Extracted records', 'Company count difference']], 'coverage_check')


if __name__ == '__main__':
    main()
