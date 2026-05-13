# Technical Documentation: Power BI RDL Filter Integration

This document explains the architecture, flow, and resolution of the RDL (Paginated Report) filter integration issues in this project.

## 1. The Core Architecture
Unlike standard Power BI reports (`.pbix`), Paginated Reports (`.rdl`) do not support client-side filtering via the Power BI JavaScript SDK. To apply custom filters, we use a hybrid approach:

1.  **XML Injection**: We modify the RDL's underlying XML to inject `<Filter>` elements into the report's DataSets.
2.  **URL Parameters**: We use the `rp:` URL prefix in the Embed URL to pass values to the parameters that drive those filters.

## 2. The Filter Flow
1.  **Discovery**: 
    - Method: `PowerBIService.DiscoverReportFilters`
    - Logic: Scans the RDL XML for `<ReportParameter>` definitions. It specifically identifies "Query-based" parameters (parameters whose values come from a DataSet).
2.  **User Interaction**: 
    - User adds a filter in the sidebar.
    - Controller calls `SaveCustomFilter`.
3.  **Application**:
    - Method: `PowerBIService.InjectManualFilterToRdl`
    - Logic: Creates a temporary version of the RDL, injects the XML `<Filter>`, and re-uploads it to Power BI.
4.  **Rendering**:
    - The frontend constructs an Embed URL: `...&rp:ParameterName=Value`.

## 3. The "Circular Dependency" Error
### The Problem
If we inject a filter into a DataSet (e.g., `CompanyList`) that is the *source* for a parameter's dropdown values, we create a circular dependency:
- The parameter needs the DataSet to populate the dropdown.
- The DataSet needs the parameter value to apply the filter.
- **Result**: "Failed to evaluate the FilterValue of the DataSet..." error.

### The Solution
1.  **DataSet Pre-Scan**: In `ApplyInjectionToXDoc`, we now identify all datasets used in `<DataSetReference>` blocks.
2.  **Intelligent Skipping**: The engine explicitly **skips** injecting filters into any dataset that is marked as a value source for a parameter.
3.  **Namespace-Agnostic Matching**: We use `LocalName` matching to ensure filters are correctly identified regardless of the RDL schema version (2010, 2016, etc.).

## 4. The "Deep Clean" Restoration
Since RDL files are permanently modified during injection, we implemented a restoration workflow:
- **Method**: `PowerBIService.CleanAndRestoreRdl`
- **Logic**: 
    1.  Strips ALL injected filters from the RDL.
    2.  Removes empty `<Filters>` tags to maintain XML schema validity.
    3.  Uploads the clean version to Power BI.
    4.  **Readiness Probe**: Polls the Power BI API to wait for the new report to be "Ready" (Active) before allowing the UI to load it. This prevents `403 Forbidden` errors.

## 5. Summary of Key Methods
- `StripAllInjections(XDocument xDoc)`: Sanitizes the XML of all system-added filters.
- `ApplyInjectionToXDoc(...)`: Safely injects new filters while avoiding circular dependencies.
- `CleanAndRestoreRdl(...)`: Orchestrates the full purge and re-upload process.
- `GetEmbedConfig(...)`: Generates the final token and URL with `rp:` parameters.

---
*Created by Antigravity AI Assistant*
