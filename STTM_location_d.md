# Source-to-Target Mapping (STTM) — `dw.location_d`

## Overview

| Attribute | Value |
|---|---|
| **Target Table** | `dw.location_d` |
| **Target Schema** | `dw` |
| **Staging Table** | `staging.location_ds` |
| **Load Strategy** | Upsert (UPDATE existing rows + INSERT new rows) |
| **Grain** | One row per `location_code` |
| **Data Filter** | `TO_TIMESTAMP('{{data_filter_end_dttm}}', 'yyyymmdd HH24:MI:SS') >= erp_start_date` |

---

## Source Systems

| Source Alias | Source Table | Schema | Description |
|---|---|---|---|
| `location_code_list` | `InventoryOrgParametersPVO` + `ffmcenter` | `oracle_erp` / `order_management_service` | UNION of ERP org codes and OMS fulfillment center names; provides the master list of location codes |
| `InventoryOrgParametersPVO` | `InventoryOrgParametersPVO` | `oracle_erp` | Oracle ERP inventory org parameters (org code, enabled flag, business unit) |
| `INV_ORG_PARAMETERS` | `INV_ORG_PARAMETERS` | `oracle_erp` | Oracle ERP inventory org attributes (warehouse type, EZ Ship flag) |
| `hr_locations` | `hr_locations` | `oracle_erp` | Oracle HR location master (address, city, state, postal code) |
| `ffmcenter` | `ffmcenter` | `order_management_service` | OMS fulfillment center master (id, name, type, address, SLA days) |
| `ffmcenter_hist` | `ffmcenter_raw` | `denver_retirement` | Historical fulfillment center data |
| `fcs` | `fulfillment_center` | `fcs` | FCS fulfillment center details (enabled flag, display name, address, SLA) |
| `tmp_location_id` | `location_d` + `codecombinationpvo` | `dw` / `oracle_erp` | Temporary table resolving `oracle_location_id` per `location_code` |

---

## Column-Level Source-to-Target Mapping

| # | Target Column | Target Data Type | Source Table(s) | Source Column(s) | Transformation / Business Logic | Nullable |
|---|---|---|---|---|---|---|
| 1 | `location_code` | VARCHAR | `location_code_list` | `location_code` | Direct — derived from UNION of `OrgOrganizationDefinitionsPOrganizationCode` (ERP) and `ffmcenter.name` (OMS) | NOT NULL |
| 2 | `location_key` | INTEGER | `dw.location_d` | `location_key` | Surrogate key: `ROW_NUMBER() OVER (ORDER BY 'location_key') + MAX(location_key)` from existing `dw.location_d`. Applied only on INSERT of new rows. | NOT NULL |
| 3 | `fulfillment_center_id` | INTEGER | `order_management_service.ffmcenter` | `id` | Direct — `ffmcenter.id`. Aggregated with `MAX()` in final SELECT. | NULLABLE |
| 4 | `fulfillment_active` | BOOLEAN | `fcs.fulfillment_center` | `fc_enabled` | Direct — `fcs.fc_enabled`. Aggregated with `MAX()`. | NULLABLE |
| 5 | `location_display_name` | VARCHAR | `fcs.fulfillment_center`, `oracle_erp.hr_locations`, `InventoryOrgParametersPVO` | `display_name`, `LegalEntityPEOName`, `region_2`, `InvOrgNamePEOName` | `NVL(LegalEntityPEOName \|\| ' (' \|\| hr_locations.region_2 \|\| ' - ' \|\| InvOrgNamePEOName \|\| ')', fcs.display_name)` — prefers ERP-derived name; falls back to FCS display name. Aggregated with `MAX()`. | NULLABLE |
| 6 | `location_active_warehouse` | INTEGER (0/1) | `oracle_erp.InventoryOrgParametersPVO` | `OrgOrganizationDefinitionsPInventoryEnabledFlag` | `CASE WHEN OrgOrganizationDefinitionsPInventoryEnabledFlag = 'Y' THEN 1 ELSE 0 END`. Aggregated with `MAX()`. | NULLABLE |
| 7 | `location_warehouse_type` | INTEGER | `oracle_erp.INV_ORG_PARAMETERS` | `ATTRIBUTE1` | `CASE WHEN ATTRIBUTE1 IN ('Retail','Pharmacy') THEN 0 WHEN ATTRIBUTE1 = 'Freezer' THEN 1 WHEN ATTRIBUTE1 = 'Cross Dock' THEN 3 END`. Aggregated with `MAX()`. | NULLABLE |
| 8 | `location_address1` | VARCHAR | `fcs.fulfillment_center`, `oracle_erp.hr_locations` | `address_line1`, `address_line_1` | `CASE WHEN TYPE = 'DROPSHIP' THEN fcs.address_line1 ELSE hr_locations.address_line_1 END`. DROPSHIP locations use FCS address; all others use HR address. Aggregated with `MAX()`. | NULLABLE |
| 9 | `location_city` | VARCHAR | `fcs.fulfillment_center`, `oracle_erp.hr_locations` | `city`, `town_or_city` | `CASE WHEN TYPE = 'DROPSHIP' THEN fcs.city ELSE hr_locations.town_or_city END`. Aggregated with `MAX()`. | NULLABLE |
| 10 | `location_post_code` | VARCHAR | `fcs.fulfillment_center`, `oracle_erp.hr_locations` | `postal_code` | `CASE WHEN TYPE = 'DROPSHIP' THEN fcs.postal_code ELSE hr_locations.postal_code END`. Aggregated with `MAX()`. | NULLABLE |
| 11 | `fulfillment_center_dropship_flag` | BOOLEAN | `order_management_service.ffmcenter` | `type` | `CASE TYPE WHEN 'DROPSHIP' THEN TRUE ELSE FALSE END`. Aggregated with `MAX()`. | NULLABLE |
| 12 | `fulfillment_center_max_pick_number` | INTEGER | `order_management_service.ffmcenter` | `maxnumpick` | Direct — `maxnumpick`. Aggregated with `MAX()`. | NULLABLE |
| 13 | `fulfillment_center_pick_delay` | INTEGER | `order_management_service.ffmcenter` | `pick_delay` | Direct — `pick_delay`. Aggregated with `MAX()`. | NULLABLE |
| 14 | `fulfillment_center_storage_rate` | NUMERIC | *(none)* | *(none)* | Hardcoded `NULL` — not currently sourced. Aggregated with `MAX()`. | NULLABLE |
| 15 | `fulfillment_center_default_shipping_offset` | INTEGER | `order_management_service.ffmcenter` | `default_ship_offset` | Direct — `default_ship_offset`. Aggregated with `MAX()`. | NULLABLE |
| 16 | `fulfillment_center_mark_for_delete` | BOOLEAN | `fcs.fulfillment_center` | `fc_enabled` | `NOT fcs.fc_enabled` — inverted active flag. Aggregated with `MAX()`. | NULLABLE |
| 17 | `legal_company_description` | VARCHAR | `location_code_list` | `location_code`, `dw_site_id` | `CASE` on `location_code` / `dw_site_id`: `'SDF1'→'Petsmart'`, `'SDF3'→'Wholesale'`, `'STK1'→'XYZ Virtual Care'`, `'STK2'→'Stark Services PLLC'`, `dw_site_id=60→'Retail Canada'`, else `'Retail'`. Aggregated with `MAX()`. | NULLABLE |
| 18 | `product_company_description` | VARCHAR | `location_code_list`, `fcs.fulfillment_center` | `location_code`, `dw_site_id`, `type` | `CASE`: `STK%→'XYZ Healthcare Services'`, `PHARMA+site 10→'XYZ Pharmacy'`, `PHARMA+site 60→'XYZ Pharmacy Canada'`, `site 60→'XYZ Canada'`, else `'XYZ'`. Aggregated with `MAX()`. | NULLABLE |
| 19 | `oracle_location_id` | INTEGER | `tmp_location_id` | `oracle_location_id` | Resolved via temp table `tmp_location_id` (see below). Joined on `location_code`. Aggregated with `MAX()`. | NULLABLE |
| 20 | `location_ez_ship_flag` | BOOLEAN | `oracle_erp.INV_ORG_PARAMETERS` | `ATTRIBUTE1` | `CASE WHEN ATTRIBUTE1 = 'EZ Ship' THEN TRUE ELSE FALSE END`. Aggregated with `MAX()`. | NULLABLE |
| 21 | `location_address2` | VARCHAR | `fcs.fulfillment_center`, `oracle_erp.hr_locations` | `address_line2`, `address_line_2` | `CASE WHEN TYPE = 'DROPSHIP' THEN fcs.address_line2 ELSE hr_locations.address_line_2 END`. Aggregated with `MAX()`. | NULLABLE |
| 22 | `location_state` | VARCHAR(2) | `fcs.fulfillment_center`, `oracle_erp.hr_locations` | `state`, `region_2` | `UPPER(LEFT(CASE WHEN TYPE = 'DROPSHIP' THEN fcs.state ELSE hr_locations.region_2 END, 2))` — upper-cased 2-character state code. Aggregated with `MAX()`. | NULLABLE |
| 23 | `location_ship_sla_days` | INTEGER | `fcs.fulfillment_center` | `ship_sla_days` | Direct — `fcs.ship_sla_days`. Aggregated with `MAX()`. | NULLABLE |

---

## Temporary Table: `tmp_location_id`

Used to resolve `oracle_location_id` for each `location_code`. Built as a local temporary table before the UPDATE/INSERT steps.

| Branch | Source | Logic |
|---|---|---|
| **Branch 1** | `dw.location_d` | Select `oracle_location_id` and `location_code` where `oracle_location_id IS NOT NULL` AND the ID does not already exist in the ERP `codecombinationpvo` join result (avoids duplicates). |
| **Branch 2** | `oracle_erp.codecombinationpvo` JOIN `dw.location_d` | Extract `codecombinationsegment3::int` as `oracle_location_id`; extract `location_code` by parsing `codecombinationdescription` with regex `-[a-zA-Z] \d-` and splitting on `-`. |

---

## Load Strategy Detail

### Step 1 — Populate Staging (`staging.location_ds`)
- Truncate/replace staging table with the full SELECT from source systems.
- Data filter: only process if `{{data_filter_end_dttm}}` ≥ ERP start date.

### Step 2 — Build `tmp_location_id`
- Resolve `oracle_location_id` per `location_code` using UNION of existing DW values and ERP code combination parsing.

### Step 3 — UPDATE existing rows in `dw.location_d`
- Match on `location_code`.
- All non-key attributes are overwritten with the latest values from staging (aggregated via `MAX()`).

### Step 4 — INSERT new rows into `dw.location_d`
- Only rows where `location_code` does NOT already exist in `dw.location_d` (LEFT JOIN + `WHERE ps.location_code IS NULL`).
- `location_key` is generated as a sequential surrogate key offset from the current maximum.

### Step 5 — COMMIT

---

## Join Diagram

```
location_code_list  (UNION: oracle_erp.InventoryOrgParametersPVO + order_management_service.ffmcenter)
        │
        ├── LEFT JOIN oracle_erp.InventoryOrgParametersPVO   ON location_code
        ├── LEFT JOIN order_management_service.ffmcenter      ON location_code = name
        ├── LEFT JOIN oracle_erp.INV_ORG_PARAMETERS           ON ORGANIZATION_ID = OrganizationId
        ├── LEFT JOIN oracle_erp.hr_locations                 ON location_code
        ├── LEFT JOIN denver_retirement.ffmcenter_raw         ON id
        └── LEFT JOIN fcs.fulfillment_center                  ON id = location_code
```

---

## Business Rules Summary

| Rule | Description |
|---|---|
| **Address source priority** | DROPSHIP type → use FCS address fields; all other types → use Oracle HR address fields |
| **Display name priority** | ERP-derived name (`LegalEntityPEOName + region_2 + InvOrgNamePEOName`) preferred; falls back to `fcs.display_name` via `NVL` |
| **Warehouse type encoding** | Retail/Pharmacy = 0, Freezer = 1, Cross Dock = 3 |
| **Active warehouse flag** | 1 if `OrgOrganizationDefinitionsPInventoryEnabledFlag = 'Y'`, else 0 |
| **Mark for delete** | Inverse of `fcs.fc_enabled` |
| **Legal company** | Specific overrides for SDF1, SDF3, STK1, STK2; Canada site (60) = 'Retail Canada'; default = 'Retail' |
| **Product company** | STK% codes → 'XYZ Healthcare Services'; PHARMA type by site; Canada site → 'XYZ Canada'; default → 'XYZ' |
| **State code** | Always upper-cased and truncated to 2 characters |
| **Storage rate** | Always NULL (not yet sourced) |
| **Surrogate key** | Generated only on INSERT; offset from current MAX to avoid collisions |
