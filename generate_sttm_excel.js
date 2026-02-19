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
    { header: 'Value',     key: 'val',  width: 70 },
  ];
  const overviewData = [
    ['Target Table',    'dw.location_d'],
    ['Target Schema',   'dw'],
    ['Staging Table',   'staging.location_ds'],
    ['Load Strategy',   'Upsert (UPDATE existing rows + INSERT new rows)'],
    ['Grain',           'One row per location_code'],
    ['Data Filter',     "TO_TIMESTAMP('{{data_filter_end_dttm}}', 'yyyymmdd HH24:MI:SS') >= erp_start_date"],
  ];
  overviewData.forEach(([attr, val]) => overviewSheet.addRow({ attr, val }));

  // Style header row
  styleHeaderRow(overviewSheet.getRow(1), '1F4E79');

  // ─── Sheet 2: Source Systems ──────────────────────────────────────────────────
  const srcSheet = workbook.addWorksheet('Source Systems');
  srcSheet.columns = [
    { header: 'Source Alias',  key: 'alias',  width: 25 },
    { header: 'Source Table',  key: 'table',  width: 45 },
    { header: 'Schema',        key: 'schema', width: 30 },
    { header: 'Description',   key: 'desc',   width: 70 },
  ];
  const srcData = [
    ['location_code_list',  'InventoryOrgParametersPVO + ffmcenter',          'oracle_erp / order_management_service', 'UNION of ERP org codes and OMS fulfillment center names; provides the master list of location codes'],
    ['InventoryOrgParametersPVO', 'InventoryOrgParametersPVO',                'oracle_erp',                            'Oracle ERP inventory org parameters (org code, enabled flag, business unit)'],
    ['INV_ORG_PARAMETERS',  'INV_ORG_PARAMETERS',                             'oracle_erp',                            'Oracle ERP inventory org attributes (warehouse type, EZ Ship flag)'],
    ['hr_locations',        'hr_locations',                                   'oracle_erp',                            'Oracle HR location master (address, city, state, postal code)'],
    ['ffmcenter',           'ffmcenter',                                      'order_management_service',              'OMS fulfillment center master (id, name, type, address, SLA days)'],
    ['ffmcenter_hist',      'ffmcenter_raw',                                  'denver_retirement',                     'Historical fulfillment center data'],
    ['fcs',                 'fulfillment_center',                             'fcs',                                   'FCS fulfillment center details (enabled flag, display name, address, SLA)'],
    ['tmp_location_id',     'location_d + codecombinationpvo',                'dw / oracle_erp',                       'Temporary table resolving oracle_location_id per location_code'],
  ];
  srcData.forEach(([alias, table, schema, desc]) => srcSheet.addRow({ alias, table, schema, desc }));
  styleHeaderRow(srcSheet.getRow(1), '1F4E79');

  // ─── Sheet 3: Column Mapping ──────────────────────────────────────────────────
  const mapSheet = workbook.addWorksheet('Column Mapping');
  mapSheet.columns = [
    { header: '#',                          key: 'num',          width: 5  },
    { header: 'Target Column',              key: 'target_col',   width: 38 },
    { header: 'Target Data Type',           key: 'target_dtype', width: 18 },
    { header: 'Source Table(s)',            key: 'src_tables',   width: 45 },
    { header: 'Source Column(s)',           key: 'src_cols',     width: 45 },
    { header: 'Transformation / Business Logic', key: 'logic',   width: 80 },
    { header: 'Mapping Type',              key: 'mapping_type', width: 30 },
    { header: 'Nullable',                  key: 'nullable',     width: 12 },
    { header: 'Notes',                     key: 'notes',        width: 60 },
  ];

  const rows = [
    {
      num: 1,
      target_col: 'location_code',
      target_dtype: 'VARCHAR',
      src_tables: 'location_code_list',
      src_cols: 'location_code',
      logic: "Direct — derived from UNION of OrgOrganizationDefinitionsPOrganizationCode (ERP) and ffmcenter.name (OMS)",
      mapping_type: 'Direct Mapping',
      nullable: 'NOT NULL',
      notes: 'Natural/business key; grain of the dimension table',
    },
    {
      num: 2,
      target_col: 'location_key',
      target_dtype: 'INTEGER',
      src_tables: 'dw.location_d',
      src_cols: 'location_key',
      logic: "Surrogate key: ROW_NUMBER() OVER (ORDER BY 'location_key') + MAX(location_key) from existing dw.location_d. Applied only on INSERT of new rows.",
      mapping_type: 'Derived/Computed',
      nullable: 'NOT NULL',
      notes: 'Surrogate key generated only on INSERT; offset from current MAX to avoid collisions',
    },
    {
      num: 3,
      target_col: 'fulfillment_center_id',
      target_dtype: 'INTEGER',
      src_tables: 'order_management_service.ffmcenter',
      src_cols: 'id',
      logic: 'Direct — ffmcenter.id. Aggregated with MAX() in final SELECT.',
      mapping_type: 'Direct Mapping',
      nullable: 'NULLABLE',
      notes: 'MAX() aggregation used to collapse multiple rows per location_code',
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
      src_tables: 'fcs.fulfillment_center, oracle_erp.hr_locations, InventoryOrgParametersPVO',
      src_cols: 'display_name, LegalEntityPEOName, region_2, InvOrgNamePEOName',
      logic: "NVL(LegalEntityPEOName || ' (' || hr_locations.region_2 || ' - ' || InvOrgNamePEOName || ')', fcs.display_name) — prefers ERP-derived name; falls back to FCS display name. Aggregated with MAX().",
      mapping_type: 'Conditional Logic/ CASE Statement',
      nullable: 'NULLABLE',
      notes: 'ERP-derived name preferred; NVL fallback to fcs.display_name',
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
      notes: "Flag encoded as integer: 1 = active, 0 = inactive",
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
      notes: 'Retail/Pharmacy=0, Freezer=1, Cross Dock=3; NULL for unrecognised values',
    },
    {
      num: 8,
      target_col: 'location_address1',
      target_dtype: 'VARCHAR',
      src_tables: 'fcs.fulfillment_center, oracle_erp.hr_locations',
      src_cols: 'address_line1, address_line_1',
      logic: 'CASE WHEN TYPE = \'DROPSHIP\' THEN fcs.address_line1 ELSE hr_locations.address_line_1 END. DROPSHIP locations use FCS address; all others use HR address. Aggregated with MAX().',
      mapping_type: 'Conditional Logic/ CASE Statement',
      nullable: 'NULLABLE',
      notes: 'Address source priority: DROPSHIP → FCS; all others → Oracle HR',
    },
    {
      num: 9,
      target_col: 'location_city',
      target_dtype: 'VARCHAR',
      src_tables: 'fcs.fulfillment_center, oracle_erp.hr_locations',
      src_cols: 'city, town_or_city',
      logic: "CASE WHEN TYPE = 'DROPSHIP' THEN fcs.city ELSE hr_locations.town_or_city END. Aggregated with MAX().",
      mapping_type: 'Conditional Logic/ CASE Statement',
      nullable: 'NULLABLE',
      notes: 'Same address source priority rule as location_address1',
    },
    {
      num: 10,
      target_col: 'location_post_code',
      target_dtype: 'VARCHAR',
      src_tables: 'fcs.fulfillment_center, oracle_erp.hr_locations',
      src_cols: 'postal_code',
      logic: "CASE WHEN TYPE = 'DROPSHIP' THEN fcs.postal_code ELSE hr_locations.postal_code END. Aggregated with MAX().",
      mapping_type: 'Conditional Logic/ CASE Statement',
      nullable: 'NULLABLE',
      notes: 'Same address source priority rule as location_address1',
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
      logic: 'Direct — maxnumpick. Aggregated with MAX().',
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
      logic: 'Direct — pick_delay. Aggregated with MAX().',
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
      logic: 'Hardcoded NULL — not currently sourced. Aggregated with MAX().',
      mapping_type: 'Hardcoded/Null',
      nullable: 'NULLABLE',
      notes: 'Placeholder column; no source data available at this time',
    },
    {
      num: 15,
      target_col: 'fulfillment_center_default_shipping_offset',
      target_dtype: 'INTEGER',
      src_tables: 'order_management_service.ffmcenter',
      src_cols: 'default_ship_offset',
      logic: 'Direct — default_ship_offset. Aggregated with MAX().',
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
      logic: 'NOT fcs.fc_enabled — inverted active flag. Aggregated with MAX().',
      mapping_type: 'Derived/Computed',
      nullable: 'NULLABLE',
      notes: 'Logical inverse of fulfillment_active; derived from the same source column fc_enabled',
    },
    {
      num: 17,
      target_col: 'legal_company_description',
      target_dtype: 'VARCHAR',
      src_tables: 'location_code_list',
      src_cols: 'location_code, dw_site_id',
      logic: "CASE on location_code / dw_site_id: 'SDF1'→'Petsmart', 'SDF3'→'Wholesale', 'STK1'→'XYZ Virtual Care', 'STK2'→'Stark Services PLLC', dw_site_id=60→'Retail Canada', else 'Retail'. Aggregated with MAX().",
      mapping_type: 'Conditional Logic/ CASE Statement',
      nullable: 'NULLABLE',
      notes: 'Specific overrides for known location codes; Canada site (dw_site_id=60) maps to Retail Canada',
    },
    {
      num: 18,
      target_col: 'product_company_description',
      target_dtype: 'VARCHAR',
      src_tables: 'location_code_list, fcs.fulfillment_center',
      src_cols: 'location_code, dw_site_id, type',
      logic: "CASE: STK%→'XYZ Healthcare Services', PHARMA+site 10→'XYZ Pharmacy', PHARMA+site 60→'XYZ Pharmacy Canada', site 60→'XYZ Canada', else 'XYZ'. Aggregated with MAX().",
      mapping_type: 'Conditional Logic/ CASE Statement',
      nullable: 'NULLABLE',
      notes: 'Multi-condition CASE using location_code pattern, dw_site_id, and fulfillment center type',
    },
    {
      num: 19,
      target_col: 'oracle_location_id',
      target_dtype: 'INTEGER',
      src_tables: 'tmp_location_id',
      src_cols: 'oracle_location_id',
      logic: 'Resolved via temp table tmp_location_id (see Temp Table sheet). Joined on location_code. Aggregated with MAX().',
      mapping_type: 'Derived/Computed',
      nullable: 'NULLABLE',
      notes: 'See "Temp Table" sheet for full resolution logic using UNION of existing DW values and ERP code combination parsing',
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
      notes: '',
    },
    {
      num: 21,
      target_col: 'location_address2',
      target_dtype: 'VARCHAR',
      src_tables: 'fcs.fulfillment_center, oracle_erp.hr_locations',
      src_cols: 'address_line2, address_line_2',
      logic: "CASE WHEN TYPE = 'DROPSHIP' THEN fcs.address_line2 ELSE hr_locations.address_line_2 END. Aggregated with MAX().",
      mapping_type: 'Conditional Logic/ CASE Statement',
      nullable: 'NULLABLE',
      notes: 'Same address source priority rule as location_address1',
    },
    {
      num: 22,
      target_col: 'location_state',
      target_dtype: 'VARCHAR(2)',
      src_tables: 'fcs.fulfillment_center, oracle_erp.hr_locations',
      src_cols: 'state, region_2',
      logic: "UPPER(LEFT(CASE WHEN TYPE = 'DROPSHIP' THEN fcs.state ELSE hr_locations.region_2 END, 2)) — upper-cased 2-character state code. Aggregated with MAX().",
      mapping_type: 'Derived/Computed',
      nullable: 'NULLABLE',
      notes: 'Always upper-cased and truncated to 2 characters regardless of source value length',
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

  // Colour-code Mapping Type column (col 7)
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
    // Wrap text for logic and notes columns
    row.getCell('logic').alignment  = { wrapText: true, vertical: 'top' };
    row.getCell('notes').alignment  = { wrapText: true, vertical: 'top' };
    row.getCell('src_tables').alignment = { wrapText: true, vertical: 'top' };
    row.getCell('src_cols').alignment   = { wrapText: true, vertical: 'top' };
  });

  // ─── Sheet 4: Temp Table ──────────────────────────────────────────────────────
  const tmpSheet = workbook.addWorksheet('Temp Table');
  tmpSheet.columns = [
    { header: 'Branch',  key: 'branch', width: 15 },
    { header: 'Source',  key: 'source', width: 35 },
    { header: 'Logic',   key: 'logic',  width: 90 },
  ];
  tmpSheet.addRow({
    branch: 'Branch 1',
    source: 'dw.location_d',
    logic:  'Select oracle_location_id and location_code where oracle_location_id IS NOT NULL AND the ID does not already exist in the ERP codecombinationpvo join result (avoids duplicates).',
  });
  tmpSheet.addRow({
    branch: 'Branch 2',
    source: 'oracle_erp.codecombinationpvo JOIN dw.location_d',
    logic:  'Extract codecombinationsegment3::int as oracle_location_id; extract location_code by parsing codecombinationdescription with regex -[a-zA-Z] \\d- and splitting on -.',
  });
  styleHeaderRow(tmpSheet.getRow(1), '1F4E79');
  tmpSheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    row.getCell('logic').alignment = { wrapText: true, vertical: 'top' };
  });

  // ─── Sheet 5: Load Strategy ───────────────────────────────────────────────────
  const loadSheet = workbook.addWorksheet('Load Strategy');
  loadSheet.columns = [
    { header: 'Step',        key: 'step',  width: 10 },
    { header: 'Name',        key: 'name',  width: 40 },
    { header: 'Description', key: 'desc',  width: 100 },
  ];
  const loadData = [
    { step: 'Step 1', name: 'Populate Staging (staging.location_ds)',  desc: 'Truncate/replace staging table with the full SELECT from source systems. Data filter: only process if {{data_filter_end_dttm}} >= ERP start date.' },
    { step: 'Step 2', name: 'Build tmp_location_id',                   desc: 'Resolve oracle_location_id per location_code using UNION of existing DW values and ERP code combination parsing.' },
    { step: 'Step 3', name: 'UPDATE existing rows in dw.location_d',   desc: 'Match on location_code. All non-key attributes are overwritten with the latest values from staging (aggregated via MAX()).' },
    { step: 'Step 4', name: 'INSERT new rows into dw.location_d',      desc: "Only rows where location_code does NOT already exist in dw.location_d (LEFT JOIN + WHERE ps.location_code IS NULL). location_key is generated as a sequential surrogate key offset from the current maximum." },
    { step: 'Step 5', name: 'COMMIT',                                  desc: 'Commit the transaction.' },
  ];
  loadData.forEach(r => loadSheet.addRow(r));
  styleHeaderRow(loadSheet.getRow(1), '1F4E79');
  loadSheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    row.getCell('desc').alignment = { wrapText: true, vertical: 'top' };
  });

  // ─── Sheet 6: Business Rules ──────────────────────────────────────────────────
  const brSheet = workbook.addWorksheet('Business Rules');
  brSheet.columns = [
    { header: 'Rule',        key: 'rule', width: 40 },
    { header: 'Description', key: 'desc', width: 100 },
  ];
  const brData = [
    { rule: 'Address source priority',    desc: 'DROPSHIP type → use FCS address fields; all other types → use Oracle HR address fields' },
    { rule: 'Display name priority',      desc: 'ERP-derived name (LegalEntityPEOName + region_2 + InvOrgNamePEOName) preferred; falls back to fcs.display_name via NVL' },
    { rule: 'Warehouse type encoding',    desc: 'Retail/Pharmacy = 0, Freezer = 1, Cross Dock = 3' },
    { rule: 'Active warehouse flag',      desc: "1 if OrgOrganizationDefinitionsPInventoryEnabledFlag = 'Y', else 0" },
    { rule: 'Mark for delete',            desc: 'Inverse of fcs.fc_enabled' },
    { rule: 'Legal company',              desc: "Specific overrides for SDF1, SDF3, STK1, STK2; Canada site (60) = 'Retail Canada'; default = 'Retail'" },
    { rule: 'Product company',            desc: "STK% codes → 'XYZ Healthcare Services'; PHARMA type by site; Canada site → 'XYZ Canada'; default → 'XYZ'" },
    { rule: 'State code',                 desc: 'Always upper-cased and truncated to 2 characters' },
    { rule: 'Storage rate',               desc: 'Always NULL (not yet sourced)' },
    { rule: 'Surrogate key',              desc: 'Generated only on INSERT; offset from current MAX to avoid collisions' },
  ];
  brData.forEach(r => brSheet.addRow(r));
  styleHeaderRow(brSheet.getRow(1), '1F4E79');
  brSheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    row.getCell('desc').alignment = { wrapText: true, vertical: 'top' };
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
