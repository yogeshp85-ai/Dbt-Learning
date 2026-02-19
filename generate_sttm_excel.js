const ExcelJS = require('exceljs');
const path = require('path');

async function generateSTTM() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'STTM Generator';
  workbook.created = new Date();

  // ─── Sheet 1: Overview ───────────────────────────────────────────────────────
  const overviewSheet = workbook.addWorksheet('Overview');
  overviewSheet.columns = [
    { header: 'Attribute', key: 'attr', width: 35 },
    { header: 'Value',     key: 'val',  width: 80 },
  ];
  const overviewData = [
    ['Target Table',    'dw.location_d'],
    ['Target Schema',   'dw'],
    ['Staging Table',   'staging.location_ds'],
    ['Load Strategy',   'Upsert (UPDATE existing rows + INSERT new rows)'],
    ['Grain',           'One row per location_code'],
    ['Data Filter',     "TO_TIMESTAMP('{{data_filter_end_dttm}}', 'yyyymmdd HH24:MI:SS') >= erp_start_date (from XYZdata.Oracle_Erp_start_Date)"],
  ];
  overviewData.forEach(([attr, val]) => overviewSheet.addRow({ attr, val }));
  styleHeaderRow(overviewSheet.getRow(1), '1F4E79');

  // ─── Sheet 2: Source Systems ──────────────────────────────────────────────────
  const srcSheet = workbook.addWorksheet('Source Systems');
  srcSheet.columns = [
    { header: 'Schema',        key: 'schema', width: 32 },
    { header: 'Table',         key: 'table',  width: 40 },
    { header: 'SQL Alias',     key: 'alias',  width: 28 },
    { header: 'Description',   key: 'desc',   width: 70 },
  ];
  const srcData = [
    { schema: 'oracle_erp',                  table: 'InventoryOrgParametersPVO',  alias: 'InventoryOrgParametersPVO / location_code_list (UNION branch 1)', desc: 'Oracle ERP inventory org parameters — provides org code (location_code), enabled flag, business unit, legal entity name' },
    { schema: 'order_management_service',    table: 'ffmcenter',                  alias: 'ffmcenter / location_code_list (UNION branch 2)',                  desc: 'OMS fulfillment center master — provides id, name (used as location_code), type, address, SLA days, pick settings' },
    { schema: 'oracle_erp',                  table: 'INV_ORG_PARAMETERS',         alias: 'INV_ORG_PARAMETERS',                                              desc: 'Oracle ERP inventory org attributes — provides warehouse type (ATTRIBUTE1) and EZ Ship flag' },
    { schema: 'oracle_erp',                  table: 'hr_locations',               alias: 'hr_locations',                                                    desc: 'Oracle HR location master — provides address, city, state/region, postal code' },
    { schema: 'denver_retirement',           table: 'ffmcenter_raw',              alias: 'ffmcenter_hist',                                                  desc: 'Historical fulfillment center data (joined but not currently used in SELECT output)' },
    { schema: 'fcs',                         table: 'fulfillment_center',         alias: 'fcs',                                                             desc: 'FCS fulfillment center details — provides enabled flag, display name, address, SLA days, type' },
    { schema: 'dw',                          table: 'location_d',                 alias: 'tmp_location_id (branch 1) / ps (UPDATE/INSERT)',                 desc: 'Target dimension table — used to resolve existing oracle_location_id and as the UPDATE/INSERT target' },
    { schema: 'oracle_erp',                  table: 'codecombinationpvo',         alias: 'erp (in tmp_location_id branch 2)',                               desc: 'ERP code combination — used to parse oracle_location_id from codecombinationdescription via regex' },
    { schema: 'XYZdata',                     table: 'Oracle_Erp_start_Date',      alias: '(subquery in WHERE clause)',                                      desc: 'Provides erp_start_date used in the data filter condition' },
  ];
  srcData.forEach(r => srcSheet.addRow(r));
  styleHeaderRow(srcSheet.getRow(1), '1F4E79');
  srcSheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    row.getCell('desc').alignment  = { wrapText: true, vertical: 'top' };
    row.getCell('alias').alignment = { wrapText: true, vertical: 'top' };
  });

  // ─── Sheet 3: Column Mapping ──────────────────────────────────────────────────
  const mapSheet = workbook.addWorksheet('Column Mapping');
  mapSheet.columns = [
    { header: '#',                               key: 'num',          width: 5  },
    { header: 'Target Column',                   key: 'target_col',   width: 38 },
    { header: 'Target Data Type',                key: 'target_dtype', width: 18 },
    { header: 'Source Schema.Table',             key: 'src_tables',   width: 50 },
    { header: 'Source Column(s)',                key: 'src_cols',     width: 45 },
    { header: 'Transformation / Business Logic', key: 'logic',        width: 85 },
    { header: 'Mapping Type',                    key: 'mapping_type', width: 32 },
    { header: 'Nullable',                        key: 'nullable',     width: 12 },
    { header: 'Notes',                           key: 'notes',        width: 65 },
  ];

  // location_code_list is a UNION subquery of:
  //   oracle_erp.InventoryOrgParametersPVO  → OrgOrganizationDefinitionsPOrganizationCode
  //   order_management_service.ffmcenter    → name
  // dw_site_id is also derived inside that subquery from businessunitpeoname

  const rows = [
    {
      num: 1,
      target_col: 'location_code',
      target_dtype: 'VARCHAR',
      src_tables: 'oracle_erp.InventoryOrgParametersPVO\norder_management_service.ffmcenter',
      src_cols: 'OrgOrganizationDefinitionsPOrganizationCode (ERP branch)\nname (OMS branch)',
      logic: "UNION of DISTINCT OrgOrganizationDefinitionsPOrganizationCode from oracle_erp.InventoryOrgParametersPVO and DISTINCT name from order_management_service.ffmcenter. The combined result forms the master list of location codes.",
      mapping_type: 'Derived/Computed',
      nullable: 'NOT NULL',
      notes: 'Natural/business key; grain of the dimension. Derived from a UNION subquery aliased as location_code_list.',
    },
    {
      num: 2,
      target_col: 'location_key',
      target_dtype: 'INTEGER',
      src_tables: 'dw.location_d',
      src_cols: 'location_key',
      logic: "ROW_NUMBER() OVER (ORDER BY 'location_key') + (SELECT MAX(location_key) FROM dw.location_d). Applied only on INSERT of new rows.",
      mapping_type: 'Derived/Computed',
      nullable: 'NOT NULL',
      notes: 'Surrogate key generated only on INSERT; offset from current MAX to avoid collisions with existing keys.',
    },
    {
      num: 3,
      target_col: 'fulfillment_center_id',
      target_dtype: 'INTEGER',
      src_tables: 'order_management_service.ffmcenter',
      src_cols: 'id',
      logic: 'Direct — ffmcenter.id. Aggregated with MAX() in UPDATE/INSERT SELECT.',
      mapping_type: 'Direct Mapping',
      nullable: 'NULLABLE',
      notes: 'MAX() aggregation used to collapse multiple staging rows per location_code.',
    },
    {
      num: 4,
      target_col: 'fulfillment_active',
      target_dtype: 'BOOLEAN',
      src_tables: 'fcs.fulfillment_center',
      src_cols: 'fc_enabled',
      logic: 'Direct — fcs.fc_enabled. Aggregated with MAX().',
      mapping_type: 'Direct Mapping',
      nullable: 'NULLABLE',
      notes: '',
    },
    {
      num: 5,
      target_col: 'location_display_name',
      target_dtype: 'VARCHAR',
      src_tables: 'oracle_erp.InventoryOrgParametersPVO\noracle_erp.hr_locations\nfcs.fulfillment_center',
      src_cols: 'LegalEntityPEOName (InventoryOrgParametersPVO)\nInvOrgNamePEOName (InventoryOrgParametersPVO)\nregion_2 (hr_locations)\ndisplay_name (fcs.fulfillment_center)',
      logic: "NVL(LegalEntityPEOName || ' (' || hr_locations.region_2 || ' - ' || InvOrgNamePEOName || ')', fcs.display_name). ERP-derived composite name preferred; falls back to fcs.display_name when ERP name is NULL. Aggregated with MAX().",
      mapping_type: 'Conditional Logic/ CASE Statement',
      nullable: 'NULLABLE',
      notes: 'NVL acts as conditional: ERP composite name takes priority; fcs.display_name is the fallback.',
    },
    {
      num: 6,
      target_col: 'location_active_warehouse',
      target_dtype: 'INTEGER (0/1)',
      src_tables: 'oracle_erp.InventoryOrgParametersPVO',
      src_cols: 'OrgOrganizationDefinitionsPInventoryEnabledFlag',
      logic: "CASE WHEN OrgOrganizationDefinitionsPInventoryEnabledFlag = 'Y' THEN 1 ELSE 0 END. Aggregated with MAX().",
      mapping_type: 'Conditional Logic/ CASE Statement',
      nullable: 'NULLABLE',
      notes: "Flag encoded as integer: 1 = active ('Y'), 0 = inactive (any other value).",
    },
    {
      num: 7,
      target_col: 'location_warehouse_type',
      target_dtype: 'INTEGER',
      src_tables: 'oracle_erp.INV_ORG_PARAMETERS',
      src_cols: 'ATTRIBUTE1',
      logic: "CASE WHEN ATTRIBUTE1 IN ('Retail','Pharmacy') THEN 0 WHEN ATTRIBUTE1 = 'Freezer' THEN 1 WHEN ATTRIBUTE1 = 'Cross Dock' THEN 3 END. Aggregated with MAX().",
      mapping_type: 'Conditional Logic/ CASE Statement',
      nullable: 'NULLABLE',
      notes: 'Retail/Pharmacy=0, Freezer=1, Cross Dock=3. NULL for unrecognised ATTRIBUTE1 values. Joined via INV_ORG_PARAMETERS.ORGANIZATION_ID = InventoryOrgParametersPVO.OrganizationId.',
    },
    {
      num: 8,
      target_col: 'location_address1',
      target_dtype: 'VARCHAR',
      src_tables: 'fcs.fulfillment_center\noracle_erp.hr_locations',
      src_cols: 'address_line1 (fcs.fulfillment_center)\naddress_line_1 (oracle_erp.hr_locations)',
      logic: "CASE WHEN TYPE = 'DROPSHIP' THEN fcs.address_line1 ELSE hr_locations.address_line_1 END. Aggregated with MAX().",
      mapping_type: 'Conditional Logic/ CASE Statement',
      nullable: 'NULLABLE',
      notes: 'TYPE comes from order_management_service.ffmcenter. DROPSHIP → FCS address; all other types → Oracle HR address.',
    },
    {
      num: 9,
      target_col: 'location_city',
      target_dtype: 'VARCHAR',
      src_tables: 'fcs.fulfillment_center\noracle_erp.hr_locations',
      src_cols: 'city (fcs.fulfillment_center)\ntown_or_city (oracle_erp.hr_locations)',
      logic: "CASE WHEN TYPE = 'DROPSHIP' THEN fcs.city ELSE hr_locations.town_or_city END. Aggregated with MAX().",
      mapping_type: 'Conditional Logic/ CASE Statement',
      nullable: 'NULLABLE',
      notes: 'Same address source priority rule as location_address1.',
    },
    {
      num: 10,
      target_col: 'location_post_code',
      target_dtype: 'VARCHAR',
      src_tables: 'fcs.fulfillment_center\noracle_erp.hr_locations',
      src_cols: 'postal_code (fcs.fulfillment_center)\npostal_code (oracle_erp.hr_locations)',
      logic: "CASE WHEN TYPE = 'DROPSHIP' THEN fcs.postal_code ELSE hr_locations.postal_code END. Aggregated with MAX().",
      mapping_type: 'Conditional Logic/ CASE Statement',
      nullable: 'NULLABLE',
      notes: 'Same address source priority rule as location_address1.',
    },
    {
      num: 11,
      target_col: 'fulfillment_center_dropship_flag',
      target_dtype: 'BOOLEAN',
      src_tables: 'order_management_service.ffmcenter',
      src_cols: 'type',
      logic: "CASE TYPE WHEN 'DROPSHIP' THEN TRUE ELSE FALSE END. Aggregated with MAX().",
      mapping_type: 'Conditional Logic/ CASE Statement',
      nullable: 'NULLABLE',
      notes: '',
    },
    {
      num: 12,
      target_col: 'fulfillment_center_max_pick_number',
      target_dtype: 'INTEGER',
      src_tables: 'order_management_service.ffmcenter',
      src_cols: 'maxnumpick',
      logic: 'Direct — ffmcenter.maxnumpick. Aggregated with MAX().',
      mapping_type: 'Direct Mapping',
      nullable: 'NULLABLE',
      notes: '',
    },
    {
      num: 13,
      target_col: 'fulfillment_center_pick_delay',
      target_dtype: 'INTEGER',
      src_tables: 'order_management_service.ffmcenter',
      src_cols: 'pick_delay',
      logic: 'Direct — ffmcenter.pick_delay. Aggregated with MAX().',
      mapping_type: 'Direct Mapping',
      nullable: 'NULLABLE',
      notes: '',
    },
    {
      num: 14,
      target_col: 'fulfillment_center_storage_rate',
      target_dtype: 'NUMERIC',
      src_tables: '(none)',
      src_cols: '(none)',
      logic: 'NULL — hardcoded literal NULL. Aggregated with MAX().',
      mapping_type: 'Hardcoded/Null',
      nullable: 'NULLABLE',
      notes: 'Placeholder column; no source data available at this time.',
    },
    {
      num: 15,
      target_col: 'fulfillment_center_default_shipping_offset',
      target_dtype: 'INTEGER',
      src_tables: 'order_management_service.ffmcenter',
      src_cols: 'default_ship_offset',
      logic: 'Direct — ffmcenter.default_ship_offset. Aggregated with MAX().',
      mapping_type: 'Direct Mapping',
      nullable: 'NULLABLE',
      notes: '',
    },
    {
      num: 16,
      target_col: 'fulfillment_center_mark_for_delete',
      target_dtype: 'BOOLEAN',
      src_tables: 'fcs.fulfillment_center',
      src_cols: 'fc_enabled',
      logic: 'NOT fcs.fc_enabled — logical NOT of the active flag. Aggregated with MAX().',
      mapping_type: 'Derived/Computed',
      nullable: 'NULLABLE',
      notes: 'Logical inverse of fulfillment_active; both columns derive from the same source column fcs.fc_enabled.',
    },
    {
      num: 17,
      target_col: 'legal_company_description',
      target_dtype: 'VARCHAR',
      src_tables: 'oracle_erp.InventoryOrgParametersPVO\norder_management_service.ffmcenter',
      src_cols: 'OrgOrganizationDefinitionsPOrganizationCode (location_code)\nbusinessunitpeoname → dw_site_id (derived in UNION subquery)',
      logic: "CASE WHEN location_code = 'SDF1' THEN 'Petsmart' WHEN location_code = 'SDF3' THEN 'Wholesale' WHEN location_code = 'STK1' THEN 'XYZ Virtual Care' WHEN location_code = 'STK2' THEN 'Stark Services PLLC' WHEN dw_site_id = 60 THEN 'Retail Canada' ELSE 'Retail' END. Aggregated with MAX().",
      mapping_type: 'Conditional Logic/ CASE Statement',
      nullable: 'NULLABLE',
      notes: "dw_site_id is derived in the UNION subquery: businessunitpeoname = 'XYZ CA BU' → 60, else 10. Specific location_code overrides take priority over site-level rule.",
    },
    {
      num: 18,
      target_col: 'product_company_description',
      target_dtype: 'VARCHAR',
      src_tables: 'oracle_erp.InventoryOrgParametersPVO\norder_management_service.ffmcenter\nfcs.fulfillment_center',
      src_cols: 'OrgOrganizationDefinitionsPOrganizationCode (location_code)\nbusinessunitpeoname → dw_site_id (derived)\ntype (fcs.fulfillment_center)',
      logic: "CASE WHEN location_code LIKE 'STK%' THEN 'XYZ Healthcare Services' WHEN fcs.type = 'PHARMA' AND dw_site_id = 10 THEN 'XYZ Pharmacy' WHEN fcs.type = 'PHARMA' AND dw_site_id = 60 THEN 'XYZ Pharmacy Canada' WHEN dw_site_id = 60 THEN 'XYZ Canada' ELSE 'XYZ' END. Aggregated with MAX().",
      mapping_type: 'Conditional Logic/ CASE Statement',
      nullable: 'NULLABLE',
      notes: 'Multi-condition CASE using location_code pattern match, dw_site_id (derived from businessunitpeoname), and fcs.type.',
    },
    {
      num: 19,
      target_col: 'oracle_location_id',
      target_dtype: 'INTEGER',
      src_tables: 'dw.location_d\noracle_erp.codecombinationpvo',
      src_cols: 'oracle_location_id (dw.location_d — branch 1)\ncodecombinationsegment3 (oracle_erp.codecombinationpvo — branch 2)\ncodecombinationdescription (oracle_erp.codecombinationpvo — branch 2)',
      logic: "Resolved via temp table tmp_location_id (UNION of two branches):\nBranch 1: SELECT DISTINCT oracle_location_id, location_code FROM dw.location_d WHERE oracle_location_id IS NOT NULL AND oracle_location_id NOT IN (codecombinationpvo derived IDs).\nBranch 2: SELECT DISTINCT codecombinationsegment3::int, split_part(regexp_substr(codecombinationdescription, '-[a-zA-Z] \\d-'), '-', 2) FROM oracle_erp.codecombinationpvo JOIN dw.location_d.\nJoined to staging on location_code. Aggregated with MAX().",
      mapping_type: 'Derived/Computed',
      nullable: 'NULLABLE',
      notes: 'See "Temp Table" sheet for full resolution logic. Branch 2 parses location_code from codecombinationdescription using regex -[a-zA-Z] \\d- and splits on -.',
    },
    {
      num: 20,
      target_col: 'location_ez_ship_flag',
      target_dtype: 'BOOLEAN',
      src_tables: 'oracle_erp.INV_ORG_PARAMETERS',
      src_cols: 'ATTRIBUTE1',
      logic: "CASE WHEN ATTRIBUTE1 = 'EZ Ship' THEN TRUE ELSE FALSE END. Aggregated with MAX().",
      mapping_type: 'Conditional Logic/ CASE Statement',
      nullable: 'NULLABLE',
      notes: "Joined via INV_ORG_PARAMETERS.ORGANIZATION_ID = InventoryOrgParametersPVO.OrganizationId.",
    },
    {
      num: 21,
      target_col: 'location_address2',
      target_dtype: 'VARCHAR',
      src_tables: 'fcs.fulfillment_center\noracle_erp.hr_locations',
      src_cols: 'address_line2 (fcs.fulfillment_center)\naddress_line_2 (oracle_erp.hr_locations)',
      logic: "CASE WHEN TYPE = 'DROPSHIP' THEN fcs.address_line2 ELSE hr_locations.address_line_2 END. Aggregated with MAX().",
      mapping_type: 'Conditional Logic/ CASE Statement',
      nullable: 'NULLABLE',
      notes: 'Same address source priority rule as location_address1.',
    },
    {
      num: 22,
      target_col: 'location_state',
      target_dtype: 'VARCHAR(2)',
      src_tables: 'fcs.fulfillment_center\noracle_erp.hr_locations',
      src_cols: 'state (fcs.fulfillment_center)\nregion_2 (oracle_erp.hr_locations)',
      logic: "UPPER(LEFT(CASE WHEN TYPE = 'DROPSHIP' THEN fcs.state ELSE hr_locations.region_2 END, 2)). Aggregated with MAX().",
      mapping_type: 'Derived/Computed',
      nullable: 'NULLABLE',
      notes: 'Always upper-cased and truncated to 2 characters. Same DROPSHIP/non-DROPSHIP address priority as other address columns.',
    },
    {
      num: 23,
      target_col: 'location_ship_sla_days',
      target_dtype: 'INTEGER',
      src_tables: 'fcs.fulfillment_center',
      src_cols: 'ship_sla_days',
      logic: 'Direct — fcs.ship_sla_days. Aggregated with MAX().',
      mapping_type: 'Direct Mapping',
      nullable: 'NULLABLE',
      notes: '',
    },
  ];

  rows.forEach(r => mapSheet.addRow(r));
  styleHeaderRow(mapSheet.getRow(1), '1F4E79');

  // Colour-code Mapping Type column
  const mappingTypeColors = {
    'Direct Mapping':                    'C6EFCE',  // green
    'Conditional Logic/ CASE Statement': 'FFEB9C',  // yellow
    'Hardcoded/Null':                    'FFCCCC',  // red-ish
    'Derived/Computed':                  'BDD7EE',  // blue
  };
  mapSheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const cell = row.getCell('mapping_type');
    const val  = cell.value;
    if (mappingTypeColors[val]) {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF' + mappingTypeColors[val] },
      };
    }
    ['logic', 'notes', 'src_tables', 'src_cols'].forEach(k => {
      row.getCell(k).alignment = { wrapText: true, vertical: 'top' };
    });
    row.getCell('target_col').alignment = { vertical: 'top' };
    row.getCell('mapping_type').alignment = { vertical: 'top' };
    row.getCell('nullable').alignment = { vertical: 'top' };
  });

  // ─── Sheet 4: Temp Table ──────────────────────────────────────────────────────
  const tmpSheet = workbook.addWorksheet('Temp Table');
  tmpSheet.columns = [
    { header: 'Branch',         key: 'branch',  width: 12 },
    { header: 'Source Tables',  key: 'source',  width: 45 },
    { header: 'Source Columns', key: 'cols',    width: 45 },
    { header: 'Logic',          key: 'logic',   width: 90 },
  ];
  tmpSheet.addRow({
    branch: 'Branch 1',
    source: 'dw.location_d',
    cols:   'oracle_location_id, location_code',
    logic:  "SELECT DISTINCT oracle_location_id::int, location_code FROM dw.location_d WHERE oracle_location_id IS NOT NULL AND oracle_location_id NOT IN (SELECT DISTINCT codecombinationsegment3::int FROM oracle_erp.codecombinationpvo JOIN dw.location_d ON split_part(regexp_substr(codecombinationdescription, '-[a-zA-Z] \\d-'), '-', 2) = location_code). Preserves existing DW values that are not superseded by ERP parsing.",
  });
  tmpSheet.addRow({
    branch: 'Branch 2',
    source: 'oracle_erp.codecombinationpvo\ndw.location_d',
    cols:   'codecombinationsegment3 → oracle_location_id\ncodecombinationdescription → location_code (via regex)',
    logic:  "SELECT DISTINCT codecombinationsegment3::int AS oracle_location_id, split_part(regexp_substr(codecombinationdescription, '-[a-zA-Z] \\d-'), '-', 2) AS location_code FROM oracle_erp.codecombinationpvo erp JOIN dw.location_d loc ON split_part(regexp_substr(erp.codecombinationdescription, '-[a-zA-Z] \\d-'), '-', 2) = loc.location_code. Extracts oracle_location_id from ERP code combinations by parsing the description field.",
  });
  styleHeaderRow(tmpSheet.getRow(1), '1F4E79');
  tmpSheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    ['logic', 'source', 'cols'].forEach(k => {
      row.getCell(k).alignment = { wrapText: true, vertical: 'top' };
    });
  });

  // ─── Sheet 5: Load Strategy ───────────────────────────────────────────────────
  const loadSheet = workbook.addWorksheet('Load Strategy');
  loadSheet.columns = [
    { header: 'Step',        key: 'step',  width: 10 },
    { header: 'Name',        key: 'name',  width: 45 },
    { header: 'Description', key: 'desc',  width: 100 },
  ];
  const loadData = [
    { step: 'Step 1', name: 'Populate Staging (staging.location_ds)',  desc: "INSERT INTO staging.location_ds: full SELECT from source systems (oracle_erp.InventoryOrgParametersPVO, order_management_service.ffmcenter, oracle_erp.INV_ORG_PARAMETERS, oracle_erp.hr_locations, denver_retirement.ffmcenter_raw, fcs.fulfillment_center). Data filter: only process if TO_TIMESTAMP('{{data_filter_end_dttm}}', 'yyyymmdd HH24:MI:SS') >= erp_start_date from XYZdata.Oracle_Erp_start_Date." },
    { step: 'Step 2', name: 'Build tmp_location_id',                   desc: 'Create local temporary table tmp_location_id (ON COMMIT PRESERVE ROWS). UNION of: (1) existing oracle_location_id values from dw.location_d not superseded by ERP parsing, and (2) oracle_location_id values parsed from oracle_erp.codecombinationpvo via regex on codecombinationdescription.' },
    { step: 'Step 3', name: 'UPDATE existing rows in dw.location_d',   desc: 'UPDATE dw.location_d SET all non-key columns from a subquery on staging.location_ds LEFT JOIN tmp_location_id, grouped by location_code with MAX() aggregation. Match condition: ps.location_code = ls.location_code.' },
    { step: 'Step 4', name: 'INSERT new rows into dw.location_d',      desc: "INSERT INTO dw.location_d: rows from staging.location_ds LEFT JOIN tmp_location_id (grouped, MAX() aggregated) LEFT JOIN dw.location_d WHERE ps.location_code IS NULL (i.e. location_code does not yet exist in target). location_key generated as ROW_NUMBER() OVER (ORDER BY 'location_key') + MAX(location_key) from dw.location_d." },
    { step: 'Step 5', name: 'COMMIT',                                  desc: 'COMMIT the transaction to persist all changes.' },
  ];
  loadData.forEach(r => loadSheet.addRow(r));
  styleHeaderRow(loadSheet.getRow(1), '1F4E79');
  loadSheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    row.getCell('desc').alignment = { wrapText: true, vertical: 'top' };
    row.getCell('name').alignment = { wrapText: true, vertical: 'top' };
  });

  // ─── Sheet 6: Business Rules ──────────────────────────────────────────────────
  const brSheet = workbook.addWorksheet('Business Rules');
  brSheet.columns = [
    { header: 'Rule',        key: 'rule', width: 40 },
    { header: 'Description', key: 'desc', width: 100 },
  ];
  const brData = [
    { rule: 'Address source priority',    desc: "DROPSHIP type (from order_management_service.ffmcenter.type) → use fcs.fulfillment_center address fields; all other types → use oracle_erp.hr_locations address fields" },
    { rule: 'Display name priority',      desc: "ERP-derived composite name (LegalEntityPEOName || ' (' || hr_locations.region_2 || ' - ' || InvOrgNamePEOName || ')') preferred; falls back to fcs.display_name via NVL when ERP name is NULL" },
    { rule: 'Warehouse type encoding',    desc: "oracle_erp.INV_ORG_PARAMETERS.ATTRIBUTE1: Retail/Pharmacy = 0, Freezer = 1, Cross Dock = 3; NULL for unrecognised values" },
    { rule: 'Active warehouse flag',      desc: "oracle_erp.InventoryOrgParametersPVO.OrgOrganizationDefinitionsPInventoryEnabledFlag: 1 if 'Y', else 0" },
    { rule: 'Mark for delete',            desc: 'Logical NOT of fcs.fulfillment_center.fc_enabled' },
    { rule: 'Legal company',              desc: "Specific overrides for SDF1→Petsmart, SDF3→Wholesale, STK1→XYZ Virtual Care, STK2→Stark Services PLLC; dw_site_id=60 (XYZ CA BU) → 'Retail Canada'; default → 'Retail'" },
    { rule: 'Product company',            desc: "STK% location codes → 'XYZ Healthcare Services'; fcs.type='PHARMA' + site 10 → 'XYZ Pharmacy'; fcs.type='PHARMA' + site 60 → 'XYZ Pharmacy Canada'; site 60 → 'XYZ Canada'; default → 'XYZ'" },
    { rule: 'State code',                 desc: 'Always UPPER-cased and truncated to 2 characters via UPPER(LEFT(..., 2))' },
    { rule: 'Storage rate',               desc: 'Always NULL — hardcoded literal; no source data available' },
    { rule: 'Surrogate key',              desc: "Generated only on INSERT; ROW_NUMBER() OVER (ORDER BY 'location_key') + MAX(location_key) from dw.location_d to avoid collisions" },
    { rule: 'dw_site_id derivation',      desc: "Derived inside the UNION subquery: businessunitpeoname = 'XYZ CA BU' → 60, else 10. Used in legal_company_description and product_company_description CASE logic." },
    { rule: 'Data filter',                desc: "Only process records where TO_TIMESTAMP('{{data_filter_end_dttm}}', 'yyyymmdd HH24:MI:SS') >= erp_start_date from XYZdata.Oracle_Erp_start_Date" },
  ];
  brData.forEach(r => brSheet.addRow(r));
  styleHeaderRow(brSheet.getRow(1), '1F4E79');
  brSheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    row.getCell('desc').alignment = { wrapText: true, vertical: 'top' };
    row.getCell('rule').alignment = { wrapText: true, vertical: 'top' };
  });

  // ─── Save ─────────────────────────────────────────────────────────────────────
  const outPath = path.join(__dirname, 'STTM_location_d.xlsx');
  await workbook.xlsx.writeFile(outPath);
  console.log('Excel file written to:', outPath);
}

function styleHeaderRow(row, argbColor) {
  row.eachCell(cell => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF' + argbColor },
    };
    cell.font  = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = {
      top:    { style: 'thin' },
      left:   { style: 'thin' },
      bottom: { style: 'thin' },
      right:  { style: 'thin' },
    };
  });
  row.height = 22;
}

generateSTTM().catch(err => { console.error(err); process.exit(1); });
