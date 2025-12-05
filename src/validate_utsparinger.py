#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Validate A4_Utsp property set on IFC utsparings models.

Validates:
1. Presence of A4_Utsp property set
2. Presence of required properties within the pset
3. Correct values/formatting of property values

Outputs:
- Excel report with validation results
- Interactive HTML report
- IFC file with NOSKI_Validering pset added
"""

import ifcopenshell
import ifcopenshell.guid
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass, field
import json
import sys
import io

from validation_rules import (
    REQUIRED_PSET,
    REQUIRED_PROPERTIES,
    DIMENSION_PROPERTIES,
    OPTIONAL_PROPERTIES,
    ALL_PROPERTIES,
    PROPERTY_VALIDATORS,
    ValidationResult,
    validate_dimensions,
    validate_id,
)

# Fix Windows console encoding
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')


@dataclass
class ElementValidation:
    """Validation results for a single IFC element."""
    guid: str
    element_name: str
    element_type: str
    object_type: Optional[str]  # IFC ObjectType attribute
    has_pset: bool
    pset_name_found: Optional[str]  # Actual pset name found (might be wrong location)
    all_psets: List[str]  # All pset names on this element
    properties_found: Dict[str, Any]
    property_validations: Dict[str, ValidationResult]
    dimension_validation: Optional[ValidationResult]
    overall_status: str  # "OK", "Feil", "Advarsel", "Mangler pset"
    error_count: int = 0
    warning_count: int = 0

    def to_dict(self) -> dict:
        """Convert to dictionary for JSON/Excel export."""
        return {
            "GUID": self.guid,
            "Element": self.element_name,
            "IFC-type": self.element_type,
            "Har A4_Utsp": "Ja" if self.has_pset else "Nei",
            "Pset funnet": self.pset_name_found or "-",
            "Status": self.overall_status,
            "Antall feil": self.error_count,
            "Antall advarsler": self.warning_count,
            **{f"{k}_verdi": str(v) if v is not None else "" for k, v in self.properties_found.items()},
            **{f"{k}_status": v.message for k, v in self.property_validations.items()},
            "Dimensjon_status": self.dimension_validation.message if self.dimension_validation else "-",
        }


@dataclass
class FileValidation:
    """Validation results for a single IFC file."""
    filename: str
    filepath: str
    total_elements: int
    elements: List[ElementValidation]
    summary: Dict[str, int] = field(default_factory=dict)

    def calculate_summary(self):
        """Calculate summary statistics."""
        self.summary = {
            "total": len(self.elements),
            "ok": sum(1 for e in self.elements if e.overall_status == "OK"),
            "feil": sum(1 for e in self.elements if e.overall_status == "Feil"),
            "advarsel": sum(1 for e in self.elements if e.overall_status == "Advarsel"),
            "mangler_pset": sum(1 for e in self.elements if e.overall_status == "Mangler pset"),
            "has_pset": sum(1 for e in self.elements if e.has_pset),
            "total_errors": sum(e.error_count for e in self.elements),
            "total_warnings": sum(e.warning_count for e in self.elements),
        }


def extract_file_prefix(filename: str) -> str:
    """Extract expected ID prefix from filename (e.g., 'A4_RIV_Utsparinger.ifc' -> 'RIV')."""
    name = Path(filename).stem
    parts = name.split('_')
    if len(parts) >= 2:
        # Try to find RIV, RIE, RIVA pattern
        for part in parts:
            if part in ['RIV', 'RIE', 'RIVA']:
                return part
    return None


def get_element_properties(element) -> Tuple[Optional[str], Dict[str, Any], List[str]]:
    """
    Extract A4_Utsp properties from an IFC element.

    Returns:
        Tuple of (pset_name_found, properties_dict, all_pset_names)
        pset_name_found is the actual pset name where properties were found
        all_pset_names is a list of all pset names on this element
    """
    properties = {}
    pset_name_found = None
    all_pset_names = []

    if not hasattr(element, 'IsDefinedBy') or not element.IsDefinedBy:
        return None, properties, all_pset_names

    for definition in element.IsDefinedBy:
        if not definition.is_a("IfcRelDefinesByProperties"):
            continue

        prop_def = definition.RelatingPropertyDefinition
        if not prop_def.is_a("IfcPropertySet"):
            continue

        pset_name = prop_def.Name
        all_pset_names.append(pset_name)

        # Check if this is the correct pset or if properties are misplaced
        if pset_name == REQUIRED_PSET:
            pset_name_found = REQUIRED_PSET
        elif any(prop.Name.startswith("A4_Utsp_") for prop in prop_def.HasProperties if hasattr(prop, 'Name')):
            # Properties found in wrong pset
            if pset_name_found is None:
                pset_name_found = f"{pset_name} (feil plassering)"

        # Extract property values
        for prop in prop_def.HasProperties:
            if not hasattr(prop, 'Name'):
                continue

            prop_name = prop.Name
            if prop_name.startswith("A4_Utsp_") or prop_name in ALL_PROPERTIES:
                value = None
                if hasattr(prop, 'NominalValue') and prop.NominalValue:
                    value = prop.NominalValue.wrappedValue
                properties[prop_name] = value

    return pset_name_found, properties, all_pset_names


def validate_element(element, expected_prefix: str = None) -> ElementValidation:
    """Validate a single IFC element against A4_Utsp requirements."""
    guid = element.GlobalId
    element_name = element.Name or "(uten navn)"
    element_type = element.is_a()
    object_type = getattr(element, 'ObjectType', None)

    pset_name_found, properties, all_psets = get_element_properties(element)
    has_pset = pset_name_found == REQUIRED_PSET

    property_validations = {}
    error_count = 0
    warning_count = 0

    # If no pset found at all
    if pset_name_found is None:
        return ElementValidation(
            guid=guid,
            element_name=element_name,
            element_type=element_type,
            object_type=object_type,
            has_pset=False,
            pset_name_found=None,
            all_psets=all_psets,
            properties_found={},
            property_validations={},
            dimension_validation=None,
            overall_status="Mangler pset",
            error_count=1,
            warning_count=0
        )

    # Validate each required property
    for prop_name in REQUIRED_PROPERTIES:
        value = properties.get(prop_name)

        if prop_name in PROPERTY_VALIDATORS:
            validator = PROPERTY_VALIDATORS[prop_name]
            # Special handling for ID validation with expected prefix
            if prop_name == "A4_Utsp_ID":
                result = validate_id(value, expected_prefix)
            else:
                result = validator(value)
        else:
            # No validator, just check presence
            if value is None or str(value).strip() == "":
                result = ValidationResult(False, "Mangler verdi", "error")
            else:
                result = ValidationResult(True, "OK")

        property_validations[prop_name] = result
        if not result.is_valid:
            if result.severity == "error":
                error_count += 1
            elif result.severity == "warning":
                warning_count += 1
        elif result.severity == "warning":
            warning_count += 1

    # Validate optional properties if present
    for prop_name in OPTIONAL_PROPERTIES:
        if prop_name in properties:
            validator = PROPERTY_VALIDATORS.get(prop_name)
            if validator:
                result = validator(properties[prop_name])
                property_validations[prop_name] = result
                if result.severity == "warning":
                    warning_count += 1

    # Validate dimensions
    dim_result = validate_dimensions(properties)
    if not dim_result.is_valid:
        error_count += 1

    # Pset in wrong location is a warning
    if pset_name_found and pset_name_found != REQUIRED_PSET:
        warning_count += 1

    # Determine overall status
    if error_count > 0:
        overall_status = "Feil"
    elif warning_count > 0:
        overall_status = "Advarsel"
    else:
        overall_status = "OK"

    return ElementValidation(
        guid=guid,
        element_name=element_name,
        element_type=element_type,
        object_type=object_type,
        has_pset=has_pset,
        pset_name_found=pset_name_found,
        all_psets=all_psets,
        properties_found=properties,
        property_validations=property_validations,
        dimension_validation=dim_result,
        overall_status=overall_status,
        error_count=error_count,
        warning_count=warning_count
    )


def validate_ifc_file(filepath: str) -> FileValidation:
    """Validate all ProvisionForVoid elements in an IFC file."""
    path = Path(filepath)
    print(f"\nValiderer: {path.name}")

    model = ifcopenshell.open(filepath)
    expected_prefix = extract_file_prefix(path.name)
    print(f"  Forventet ID-prefiks: {expected_prefix or '(ikke detektert)'}")

    # Get all IfcBuildingElementProxy (typical for ProvisionForVoid)
    elements = list(model.by_type("IfcBuildingElementProxy"))
    print(f"  Antall elementer: {len(elements)}")

    validated_elements = []
    for element in elements:
        validation = validate_element(element, expected_prefix)
        validated_elements.append(validation)

    file_validation = FileValidation(
        filename=path.name,
        filepath=str(path),
        total_elements=len(elements),
        elements=validated_elements
    )
    file_validation.calculate_summary()

    # Print summary
    s = file_validation.summary
    print(f"  Resultater:")
    print(f"    OK: {s['ok']}")
    print(f"    Feil: {s['feil']}")
    print(f"    Advarsler: {s['advarsel']}")
    print(f"    Mangler pset: {s['mangler_pset']}")

    return file_validation


def generate_excel_report(validations: List[FileValidation], output_path: str):
    """Generate Excel report with validation results."""
    import pandas as pd

    print(f"\nGenererer Excel-rapport: {output_path}")

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Summary sheet
        summary_data = []
        for fv in validations:
            s = fv.summary
            summary_data.append({
                "Fil": fv.filename,
                "Totalt": s['total'],
                "OK": s['ok'],
                "Feil": s['feil'],
                "Advarsler": s['advarsel'],
                "Mangler pset": s['mangler_pset'],
                "Har A4_Utsp": s['has_pset'],
                "Dekningsgrad A4_Utsp": f"{s['has_pset'] / s['total'] * 100:.1f}%" if s['total'] > 0 else "0%",
                "Totalt antall feil": s['total_errors'],
                "Totalt antall advarsler": s['total_warnings'],
            })

        df_summary = pd.DataFrame(summary_data)
        df_summary.to_excel(writer, sheet_name="Oversikt", index=False)

        # Detail sheets per file
        for fv in validations:
            sheet_name = fv.filename[:28]  # Excel sheet name limit
            rows = []
            for ev in fv.elements:
                # Collect error and warning messages for summary
                error_messages = []
                warning_messages = []

                for prop_name, result in ev.property_validations.items():
                    short_name = prop_name.replace('A4_Utsp_', '')
                    if not result.is_valid:
                        if result.severity == "error":
                            error_messages.append(f"{short_name}: {result.message}")
                        elif result.severity == "warning":
                            warning_messages.append(f"{short_name}: {result.message}")
                    elif result.severity == "warning":
                        warning_messages.append(f"{short_name}: {result.message}")

                if ev.dimension_validation and not ev.dimension_validation.is_valid:
                    error_messages.append(f"Dimensjoner: {ev.dimension_validation.message}")

                # For missing pset, provide useful context
                if not ev.has_pset and ev.pset_name_found is None:
                    if ev.all_psets:
                        error_messages.append(f"Mangler A4_Utsp pset. Funnet psets: {', '.join(ev.all_psets)}")
                    else:
                        error_messages.append(f"Mangler A4_Utsp pset. Ingen psets funnet - sjekk GUID {ev.guid} i IFC")
                elif ev.pset_name_found and ev.pset_name_found != "A4_Utsp":
                    warning_messages.append(f"A4_Utsp-egenskaper funnet i feil pset: {ev.pset_name_found}")

                # Combine messages
                all_messages = []
                if error_messages:
                    all_messages.extend([f"FEIL: {m}" for m in error_messages])
                if warning_messages:
                    all_messages.extend([f"ADVARSEL: {m}" for m in warning_messages])

                row = {
                    "GUID": ev.guid,
                    "Navn": ev.element_name,
                    "ObjectType": ev.object_type or "-",
                    "IFC-type": ev.element_type,
                    "Status": ev.overall_status,
                    "Feilmeldinger": "; ".join(all_messages) if all_messages else "-",
                    "Har A4_Utsp": "Ja" if ev.has_pset else "Nei",
                    "Pset funnet": ev.pset_name_found or "-",
                    "Tilgjengelige psets": ", ".join(ev.all_psets) if ev.all_psets else "-",
                    "Antall feil": ev.error_count,
                    "Antall advarsler": ev.warning_count,
                }

                # Add property values and validation status
                for prop_name in REQUIRED_PROPERTIES + DIMENSION_PROPERTIES:
                    row[prop_name] = ev.properties_found.get(prop_name, "")
                    if prop_name in ev.property_validations:
                        row[f"{prop_name}_sjekk"] = ev.property_validations[prop_name].message
                    else:
                        row[f"{prop_name}_sjekk"] = "-"

                # Dimension check
                if ev.dimension_validation:
                    row["Dimensjon_sjekk"] = ev.dimension_validation.message
                else:
                    row["Dimensjon_sjekk"] = "-"

                rows.append(row)

            df_detail = pd.DataFrame(rows)
            df_detail.to_excel(writer, sheet_name=sheet_name, index=False)

        # Error list sheet
        error_rows = []
        for fv in validations:
            for ev in fv.elements:
                if ev.error_count > 0 or ev.warning_count > 0:
                    for prop_name, result in ev.property_validations.items():
                        if not result.is_valid or result.severity == "warning":
                            error_rows.append({
                                "Fil": fv.filename,
                                "GUID": ev.guid,
                                "Navn": ev.element_name,
                                "Egenskap": prop_name,
                                "Verdi": ev.properties_found.get(prop_name, ""),
                                "Melding": result.message,
                                "Alvorlighet": "Feil" if result.severity == "error" else "Advarsel",
                            })
                    if ev.dimension_validation and not ev.dimension_validation.is_valid:
                        error_rows.append({
                            "Fil": fv.filename,
                            "GUID": ev.guid,
                            "Navn": ev.element_name,
                            "Egenskap": "Dimensjoner",
                            "Verdi": "-",
                            "Melding": ev.dimension_validation.message,
                            "Alvorlighet": "Feil",
                        })

        if error_rows:
            df_errors = pd.DataFrame(error_rows)
            df_errors.to_excel(writer, sheet_name="Feil og advarsler", index=False)

    print(f"  Lagret: {output_path}")


def generate_html_report(validations: List[FileValidation], output_path: str):
    """Generate interactive HTML report with validation results."""
    print(f"\nGenererer HTML-rapport: {output_path}")

    # Prepare data for JavaScript
    js_data = {
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "files": []
    }

    for fv in validations:
        file_data = {
            "filename": fv.filename,
            "summary": fv.summary,
            "elements": []
        }
        for ev in fv.elements:
            # Collect error and warning messages
            error_list = []
            warning_list = []

            for prop_name, result in ev.property_validations.items():
                short_name = prop_name.replace('A4_Utsp_', '')
                if not result.is_valid:
                    if result.severity == "error":
                        error_list.append(f"{short_name}: {result.message}")
                    elif result.severity == "warning":
                        warning_list.append(f"{short_name}: {result.message}")
                elif result.severity == "warning":
                    warning_list.append(f"{short_name}: {result.message}")

            if ev.dimension_validation and not ev.dimension_validation.is_valid:
                error_list.append(f"Dimensjoner: {ev.dimension_validation.message}")

            # For missing pset, provide useful context
            if not ev.has_pset and ev.pset_name_found is None:
                if ev.all_psets:
                    error_list.append(f"Mangler A4_Utsp. Har: {', '.join(ev.all_psets)}")
                else:
                    error_list.append(f"Mangler A4_Utsp. Sjekk GUID i IFC")
            elif ev.pset_name_found and ev.pset_name_found != "A4_Utsp":
                warning_list.append(f"Egenskaper i feil pset: {ev.pset_name_found}")

            # Combine for display
            all_messages = []
            all_messages.extend([f"FEIL: {m}" for m in error_list])
            all_messages.extend([f"ADVARSEL: {m}" for m in warning_list])

            elem_data = {
                "guid": ev.guid,
                "name": ev.element_name,
                "type": ev.element_type,
                "objectType": ev.object_type or "-",
                "status": ev.overall_status,
                "hasPset": ev.has_pset,
                "psetFound": ev.pset_name_found,
                "allPsets": ev.all_psets,
                "errorCount": ev.error_count,
                "warningCount": ev.warning_count,
                "errorMessages": all_messages,
                "properties": {k: str(v) if v is not None else "" for k, v in ev.properties_found.items()},
                "validations": {k: {"valid": v.is_valid, "message": v.message, "severity": v.severity}
                               for k, v in ev.property_validations.items()},
            }
            if ev.dimension_validation:
                elem_data["dimensionValidation"] = {
                    "valid": ev.dimension_validation.is_valid,
                    "message": ev.dimension_validation.message,
                    "severity": ev.dimension_validation.severity
                }
            file_data["elements"].append(elem_data)
        js_data["files"].append(file_data)

    html_content = f'''<!DOCTYPE html>
<html lang="no">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>A4_Utsp Valideringsrapport</title>
    <style>
        * {{ box-sizing: border-box; margin: 0; padding: 0; }}
        body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #f5f5f5; color: #333; }}
        .container {{ max-width: 1400px; margin: 0 auto; padding: 20px; }}
        header {{ background: linear-gradient(135deg, #1a365d 0%, #2c5282 100%); color: white; padding: 30px; border-radius: 12px; margin-bottom: 20px; }}
        header h1 {{ font-size: 1.8rem; margin-bottom: 8px; }}
        header p {{ opacity: 0.9; font-size: 0.95rem; }}
        .summary-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-bottom: 20px; }}
        .summary-card {{ background: white; border-radius: 10px; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }}
        .summary-card h3 {{ font-size: 0.85rem; color: #666; margin-bottom: 8px; text-transform: uppercase; letter-spacing: 0.5px; }}
        .summary-card .value {{ font-size: 2rem; font-weight: 700; }}
        .summary-card.ok .value {{ color: #38a169; }}
        .summary-card.error .value {{ color: #e53e3e; }}
        .summary-card.warning .value {{ color: #dd6b20; }}
        .summary-card.missing .value {{ color: #805ad5; }}
        .file-section {{ background: white; border-radius: 12px; padding: 25px; margin-bottom: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }}
        .file-section.collapsed .file-content {{ display: none; }}
        .file-section.collapsed .file-header {{ margin-bottom: 0; padding-bottom: 0; border-bottom: none; }}
        .file-header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; padding-bottom: 15px; border-bottom: 2px solid #e2e8f0; cursor: pointer; }}
        .file-header:hover {{ background: #f7fafc; margin: -10px; padding: 10px; padding-bottom: 25px; border-radius: 8px; }}
        .file-header h2 {{ font-size: 1.3rem; color: #1a365d; display: flex; align-items: center; gap: 10px; }}
        .toggle-icon {{ font-size: 0.8rem; transition: transform 0.2s; color: #718096; }}
        .file-section.collapsed .toggle-icon {{ transform: rotate(-90deg); }}
        .file-stats {{ display: flex; gap: 15px; }}
        .file-stat {{ padding: 5px 12px; border-radius: 20px; font-size: 0.85rem; font-weight: 600; }}
        .file-stat.ok {{ background: #c6f6d5; color: #276749; }}
        .file-stat.error {{ background: #fed7d7; color: #c53030; }}
        .file-stat.warning {{ background: #feebc8; color: #c05621; }}
        .file-stat.missing {{ background: #e9d8fd; color: #6b46c1; }}
        .filters {{ display: flex; gap: 10px; margin-bottom: 20px; flex-wrap: wrap; }}
        .filter-btn {{ padding: 8px 16px; border: 2px solid #e2e8f0; border-radius: 8px; background: white; cursor: pointer; font-size: 0.9rem; transition: all 0.2s; }}
        .filter-btn:hover {{ border-color: #4299e1; }}
        .filter-btn.active {{ background: #4299e1; color: white; border-color: #4299e1; }}
        .search-box {{ padding: 10px 15px; border: 2px solid #e2e8f0; border-radius: 8px; font-size: 0.9rem; width: 250px; }}
        .search-box:focus {{ outline: none; border-color: #4299e1; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 0.9rem; }}
        th {{ background: #f7fafc; padding: 12px 10px; text-align: left; font-weight: 600; color: #4a5568; border-bottom: 2px solid #e2e8f0; position: sticky; top: 0; }}
        td {{ padding: 10px; border-bottom: 1px solid #e2e8f0; }}
        tr:hover {{ background: #f7fafc; }}
        .status-badge {{ display: inline-block; padding: 4px 10px; border-radius: 12px; font-size: 0.8rem; font-weight: 600; }}
        .status-badge.ok {{ background: #c6f6d5; color: #276749; }}
        .status-badge.feil {{ background: #fed7d7; color: #c53030; }}
        .status-badge.advarsel {{ background: #feebc8; color: #c05621; }}
        .status-badge.mangler {{ background: #e9d8fd; color: #6b46c1; }}
        .expandable {{ cursor: pointer; }}
        .expandable:hover {{ background: #edf2f7; }}
        .details-row {{ display: none; }}
        .details-row.show {{ display: table-row; }}
        .details-content {{ padding: 20px; background: #f7fafc; }}
        .prop-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 10px; }}
        .prop-item {{ display: flex; justify-content: space-between; padding: 8px 12px; background: white; border-radius: 6px; border-left: 3px solid #e2e8f0; }}
        .prop-item.valid {{ border-left-color: #38a169; }}
        .prop-item.invalid {{ border-left-color: #e53e3e; }}
        .prop-item.warning {{ border-left-color: #dd6b20; }}
        .prop-name {{ font-weight: 500; color: #4a5568; }}
        .prop-value {{ color: #718096; max-width: 150px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
        .guid {{ font-family: monospace; font-size: 0.85rem; color: #718096; }}
        .error-msg {{ font-size: 0.85rem; color: #c53030; max-width: 300px; }}
        .toggle-all-container {{ display: flex; gap: 10px; margin-bottom: 15px; }}
        .toggle-all-btn {{ padding: 8px 16px; border: 2px solid #e2e8f0; border-radius: 8px; background: white; cursor: pointer; font-size: 0.85rem; transition: all 0.2s; }}
        .toggle-all-btn:hover {{ border-color: #4299e1; background: #ebf8ff; }}
        .footer {{ text-align: center; padding: 20px; color: #718096; font-size: 0.85rem; }}
        .progress-bar {{ height: 8px; background: #e2e8f0; border-radius: 4px; overflow: hidden; margin-top: 10px; }}
        .progress-fill {{ height: 100%; transition: width 0.3s; }}
        .progress-fill.ok {{ background: #38a169; }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>A4_Utsp Valideringsrapport</h1>
            <p>Generert: {js_data["generated"]} | Utsparingsvalidering iht. BIM-gjennomføringsplan</p>
        </header>

        <div id="summary-section"></div>
        <div class="toggle-all-container">
            <button class="toggle-all-btn" onclick="collapseAll()">Skjul alle</button>
            <button class="toggle-all-btn" onclick="expandAll()">Vis alle</button>
        </div>
        <div id="files-section"></div>

        <div class="footer">
            <p>Valideringsrapport generert av pset-validation</p>
        </div>
    </div>

    <script>
        const data = {json.dumps(js_data, ensure_ascii=False)};

        function renderSummary() {{
            const totals = data.files.reduce((acc, f) => {{
                acc.total += f.summary.total;
                acc.ok += f.summary.ok;
                acc.feil += f.summary.feil;
                acc.advarsel += f.summary.advarsel;
                acc.mangler_pset += f.summary.mangler_pset;
                return acc;
            }}, {{ total: 0, ok: 0, feil: 0, advarsel: 0, mangler_pset: 0 }});

            const pct = totals.total > 0 ? ((totals.ok / totals.total) * 100).toFixed(1) : 0;

            document.getElementById('summary-section').innerHTML = `
                <div class="summary-grid">
                    <div class="summary-card">
                        <h3>Totalt elementer</h3>
                        <div class="value">${{totals.total}}</div>
                    </div>
                    <div class="summary-card ok">
                        <h3>OK</h3>
                        <div class="value">${{totals.ok}}</div>
                    </div>
                    <div class="summary-card error">
                        <h3>Feil</h3>
                        <div class="value">${{totals.feil}}</div>
                    </div>
                    <div class="summary-card warning">
                        <h3>Advarsler</h3>
                        <div class="value">${{totals.advarsel}}</div>
                    </div>
                    <div class="summary-card missing">
                        <h3>Mangler pset</h3>
                        <div class="value">${{totals.mangler_pset}}</div>
                    </div>
                    <div class="summary-card">
                        <h3>Godkjenningsgrad</h3>
                        <div class="value">${{pct}}%</div>
                        <div class="progress-bar"><div class="progress-fill ok" style="width: ${{pct}}%"></div></div>
                    </div>
                </div>
            `;
        }}

        function renderFiles() {{
            let html = '';
            data.files.forEach((file, fileIdx) => {{
                html += `
                    <div class="file-section" id="file-${{fileIdx}}">
                        <div class="file-header" onclick="toggleFileSection(${{fileIdx}})">
                            <h2><span class="toggle-icon">▼</span>${{file.filename}}</h2>
                            <div class="file-stats">
                                <span class="file-stat ok">${{file.summary.ok}} OK</span>
                                <span class="file-stat error">${{file.summary.feil}} Feil</span>
                                <span class="file-stat warning">${{file.summary.advarsel}} Advarsler</span>
                                <span class="file-stat missing">${{file.summary.mangler_pset}} Mangler</span>
                            </div>
                        </div>
                        <div class="file-content">
                            <div class="filters">
                                <button class="filter-btn active" onclick="event.stopPropagation(); filterElements(${{fileIdx}}, 'all')">Alle</button>
                                <button class="filter-btn" onclick="event.stopPropagation(); filterElements(${{fileIdx}}, 'ok')">OK</button>
                                <button class="filter-btn" onclick="event.stopPropagation(); filterElements(${{fileIdx}}, 'feil')">Feil</button>
                                <button class="filter-btn" onclick="event.stopPropagation(); filterElements(${{fileIdx}}, 'advarsel')">Advarsler</button>
                                <button class="filter-btn" onclick="event.stopPropagation(); filterElements(${{fileIdx}}, 'mangler')">Mangler pset</button>
                                <input type="text" class="search-box" placeholder="Søk etter ID, GUID, navn..." onkeyup="searchElements(${{fileIdx}}, this.value)" onclick="event.stopPropagation()">
                            </div>
                            <table id="table-${{fileIdx}}">
                                <thead>
                                    <tr>
                                        <th>Status</th>
                                        <th>A4_Utsp_ID</th>
                                        <th>ObjectType</th>
                                        <th>Navn</th>
                                        <th>Feilmeldinger</th>
                                        <th>GUID</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    ${{file.elements.map((elem, elemIdx) => renderElementRow(elem, fileIdx, elemIdx)).join('')}}
                                </tbody>
                            </table>
                        </div>
                    </div>
                `;
            }});
            document.getElementById('files-section').innerHTML = html;
        }}

        function toggleFileSection(fileIdx) {{
            const section = document.getElementById(`file-${{fileIdx}}`);
            section.classList.toggle('collapsed');
        }}

        function collapseAll() {{
            document.querySelectorAll('.file-section').forEach(s => s.classList.add('collapsed'));
        }}

        function expandAll() {{
            document.querySelectorAll('.file-section').forEach(s => s.classList.remove('collapsed'));
        }}

        function renderElementRow(elem, fileIdx, elemIdx) {{
            const statusClass = elem.status.toLowerCase().replace(' ', '-');
            const id = elem.properties['A4_Utsp_ID'] || '-';
            const errors = elem.errorMessages.length > 0 ? elem.errorMessages.join('; ') : '-';
            return `
                <tr class="expandable element-row" data-status="${{statusClass}}" data-search="${{id}} ${{elem.guid}} ${{elem.name}} ${{elem.objectType}}".toLowerCase() onclick="toggleDetails(${{fileIdx}}, ${{elemIdx}})">
                    <td><span class="status-badge ${{statusClass}}">${{elem.status}}</span></td>
                    <td>${{id}}</td>
                    <td>${{elem.objectType}}</td>
                    <td>${{elem.name}}</td>
                    <td class="error-msg">${{errors}}</td>
                    <td class="guid">${{elem.guid}}</td>
                </tr>
                <tr class="details-row" id="details-${{fileIdx}}-${{elemIdx}}">
                    <td colspan="6">
                        <div class="details-content">
                            <div class="prop-grid">
                                ${{Object.entries(elem.properties).map(([k, v]) => {{
                                    const validation = elem.validations[k];
                                    let cls = '';
                                    let msg = '';
                                    if (validation) {{
                                        cls = validation.valid ? 'valid' : (validation.severity === 'warning' ? 'warning' : 'invalid');
                                        msg = validation.message;
                                    }}
                                    return `<div class="prop-item ${{cls}}" title="${{msg}}">
                                        <span class="prop-name">${{k.replace('A4_Utsp_', '')}}</span>
                                        <span class="prop-value">${{v || '(tom)'}}</span>
                                    </div>`;
                                }}).join('')}}
                                ${{elem.dimensionValidation ? `
                                    <div class="prop-item ${{elem.dimensionValidation.valid ? 'valid' : 'invalid'}}" title="${{elem.dimensionValidation.message}}">
                                        <span class="prop-name">Dimensjoner</span>
                                        <span class="prop-value">${{elem.dimensionValidation.message}}</span>
                                    </div>
                                ` : ''}}
                            </div>
                        </div>
                    </td>
                </tr>
            `;
        }}

        function toggleDetails(fileIdx, elemIdx) {{
            const row = document.getElementById(`details-${{fileIdx}}-${{elemIdx}}`);
            row.classList.toggle('show');
        }}

        function filterElements(fileIdx, status) {{
            const table = document.getElementById(`table-${{fileIdx}}`);
            const rows = table.querySelectorAll('.element-row');
            const buttons = table.parentElement.querySelectorAll('.filter-btn');

            buttons.forEach(btn => btn.classList.remove('active'));
            event.target.classList.add('active');

            rows.forEach(row => {{
                const detailsId = row.getAttribute('onclick').match(/\\d+, (\\d+)/);
                const detailsRow = document.getElementById(`details-${{fileIdx}}-${{detailsId[1]}}`);

                if (status === 'all' || row.dataset.status === status ||
                    (status === 'mangler' && row.dataset.status === 'mangler-pset')) {{
                    row.style.display = '';
                }} else {{
                    row.style.display = 'none';
                    detailsRow.classList.remove('show');
                }}
            }});
        }}

        function searchElements(fileIdx, query) {{
            const table = document.getElementById(`table-${{fileIdx}}`);
            const rows = table.querySelectorAll('.element-row');
            const q = query.toLowerCase();

            rows.forEach(row => {{
                const text = row.dataset.search;
                row.style.display = text.includes(q) ? '' : 'none';
            }});
        }}

        renderSummary();
        renderFiles();
    </script>
</body>
</html>'''

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f"  Lagret: {output_path}")


def add_validation_pset_to_ifc(validation: FileValidation, output_path: str):
    """Add NOSKI_Validering property set to IFC file."""
    print(f"\nLegger til NOSKI_Validering pset: {output_path}")

    model = ifcopenshell.open(validation.filepath)

    # Create lookup for validation results
    validation_lookup = {ev.guid: ev for ev in validation.elements}

    updated_count = 0
    for element in model:
        if not hasattr(element, 'GlobalId'):
            continue

        guid = element.GlobalId
        if guid not in validation_lookup:
            continue

        ev = validation_lookup[guid]

        # Determine validation status text
        if ev.overall_status == "OK":
            val_status = "Godkjent"
        elif ev.overall_status == "Advarsel":
            val_status = "Godkjent med advarsler"
        elif ev.overall_status == "Mangler pset":
            val_status = "Mangler A4_Utsp"
        else:
            val_status = "Ikke godkjent"

        # Pset status
        if ev.has_pset:
            pset_status = "OK"
        elif ev.pset_name_found:
            pset_status = f"Feil plassering: {ev.pset_name_found}"
        else:
            pset_status = "Mangler"

        # Dimension status
        dim_status = ev.dimension_validation.message if ev.dimension_validation else "Ikke validert"

        # Collect error and warning messages
        error_messages = []
        warning_messages = []

        for prop_name, result in ev.property_validations.items():
            short_name = prop_name.replace('A4_Utsp_', '')
            if not result.is_valid:
                if result.severity == "error":
                    error_messages.append(f"{short_name}: {result.message}")
                elif result.severity == "warning":
                    warning_messages.append(f"{short_name}: {result.message}")
            elif result.severity == "warning":
                warning_messages.append(f"{short_name}: {result.message}")

        if ev.dimension_validation and not ev.dimension_validation.is_valid:
            error_messages.append(f"Dimensjoner: {ev.dimension_validation.message}")

        # For missing pset, provide useful context
        if not ev.has_pset and ev.pset_name_found is None:
            if ev.all_psets:
                error_messages.append(f"Mangler A4_Utsp. Har: {', '.join(ev.all_psets)}")
            else:
                error_messages.append(f"Mangler A4_Utsp. Sjekk GUID i IFC")
        elif ev.pset_name_found and ev.pset_name_found != "A4_Utsp":
            warning_messages.append(f"Egenskaper i feil pset: {ev.pset_name_found}")

        # Combine messages
        all_messages = []
        all_messages.extend([f"FEIL: {m}" for m in error_messages])
        all_messages.extend([f"ADVARSEL: {m}" for m in warning_messages])
        error_summary = "; ".join(all_messages) if all_messages else "-"

        # Create properties
        property_values = [
            model.create_entity("IfcPropertySingleValue",
                               Name="Valideringsstatus",
                               NominalValue=model.create_entity("IfcLabel", val_status)),
            model.create_entity("IfcPropertySingleValue",
                               Name="ObjectType",
                               NominalValue=model.create_entity("IfcLabel", ev.object_type or "-")),
            model.create_entity("IfcPropertySingleValue",
                               Name="Feilmeldinger",
                               NominalValue=model.create_entity("IfcText", error_summary)),
            model.create_entity("IfcPropertySingleValue",
                               Name="Antall feil",
                               NominalValue=model.create_entity("IfcInteger", ev.error_count)),
            model.create_entity("IfcPropertySingleValue",
                               Name="Antall advarsler",
                               NominalValue=model.create_entity("IfcInteger", ev.warning_count)),
            model.create_entity("IfcPropertySingleValue",
                               Name="A4_Utsp pset",
                               NominalValue=model.create_entity("IfcLabel", pset_status)),
            model.create_entity("IfcPropertySingleValue",
                               Name="Tilgjengelige psets",
                               NominalValue=model.create_entity("IfcLabel", ", ".join(ev.all_psets) if ev.all_psets else "-")),
            model.create_entity("IfcPropertySingleValue",
                               Name="Dimensjoner",
                               NominalValue=model.create_entity("IfcLabel", dim_status)),
        ]

        # Add individual property validations
        for prop_name, result in ev.property_validations.items():
            short_name = prop_name.replace("A4_Utsp_", "")
            property_values.append(
                model.create_entity("IfcPropertySingleValue",
                                   Name=f"{short_name}_sjekk",
                                   NominalValue=model.create_entity("IfcLabel", result.message))
            )

        # Create property set
        property_set = model.create_entity("IfcPropertySet",
                                           GlobalId=ifcopenshell.guid.new(),
                                           Name="NOSKI_Validering",
                                           HasProperties=property_values)

        model.create_entity("IfcRelDefinesByProperties",
                           GlobalId=ifcopenshell.guid.new(),
                           RelatedObjects=[element],
                           RelatingPropertyDefinition=property_set)

        updated_count += 1

    model.write(output_path)
    print(f"  Oppdatert {updated_count} elementer")
    print(f"  Lagret: {output_path}")


def main():
    """Main entry point."""
    import argparse

    parser = argparse.ArgumentParser(description='Valider A4_Utsp pset på IFC utsparingsmodeller')
    parser.add_argument('input', nargs='+', help='IFC-fil(er) eller mappe med IFC-filer')
    parser.add_argument('-o', '--output', default='data/output', help='Utmappe for rapporter')
    parser.add_argument('--no-excel', action='store_true', help='Ikke generer Excel-rapport')
    parser.add_argument('--no-html', action='store_true', help='Ikke generer HTML-rapport')
    parser.add_argument('--no-ifc', action='store_true', help='Ikke generer IFC med NOSKI_Validering')

    args = parser.parse_args()

    # Collect IFC files
    ifc_files = []
    for path_str in args.input:
        path = Path(path_str)
        if path.is_dir():
            ifc_files.extend(path.glob('*.ifc'))
        elif path.suffix.lower() == '.ifc':
            ifc_files.append(path)

    if not ifc_files:
        print("Ingen IFC-filer funnet!")
        return

    print("=" * 70)
    print("A4_Utsp Validering")
    print("=" * 70)
    print(f"Filer å validere: {len(ifc_files)}")
    for f in ifc_files:
        print(f"  - {f.name}")

    # Validate each file
    validations = []
    for ifc_file in ifc_files:
        validation = validate_ifc_file(str(ifc_file))
        validations.append(validation)

    # Create output directory
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M")

    # Generate reports
    if not args.no_excel:
        excel_path = output_dir / f"A4_Utsp_validering_{timestamp}.xlsx"
        generate_excel_report(validations, str(excel_path))

    if not args.no_html:
        html_path = output_dir / f"A4_Utsp_validering_{timestamp}.html"
        generate_html_report(validations, str(html_path))

    if not args.no_ifc:
        for validation in validations:
            ifc_name = Path(validation.filename).stem
            ifc_output = output_dir / f"{ifc_name}_validert.ifc"
            add_validation_pset_to_ifc(validation, str(ifc_output))

    print("\n" + "=" * 70)
    print("Validering fullført!")
    print("=" * 70)


if __name__ == "__main__":
    main()
