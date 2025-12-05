#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
A4_Utsp Validering - Streamlit App

Upload IFC utsparingsmodeller for validering mot BIM-gjennomføringsplan krav.
"""

import streamlit as st
import ifcopenshell
import ifcopenshell.guid
import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass, field
import tempfile
import io
import re
import json

# ============================================================================
# Validation Rules (inline for single-file deployment)
# ============================================================================

REQUIRED_PSET = "A4_Utsp"

REQUIRED_PROPERTIES = [
    "A4_Utsp_Kategori",
    "A4_Utsp_ID",
    "A4_Utsp_Utsparingstype",
    "A4_Utsp_Tetting",
    "A4_Utsp_Fase",
    "A4_Utsp_Status",
    "A4_Utsp_Rev",
    "A4_Utsp_RevDato",
    "A4_Utsp_RevBeskrivelse",
]

DIMENSION_PROPERTIES = [
    "A4_Utsp_DimBredde",
    "A4_Utsp_DimHøyde",
    "A4_Utsp_DimDybde",
    "A4_Utsp_DimDiameter",
]

OPTIONAL_PROPERTIES = ["A4_Utsp_Funksjon"]
ALL_PROPERTIES = REQUIRED_PROPERTIES + DIMENSION_PROPERTIES + OPTIONAL_PROPERTIES

# Valid values
VALID_KATEGORI = ["ProvisionForVoid"]
VALID_UTSPARINGSTYPE = ["Utsparing", "Hulltaking", "Innstøpningsgods"]
VALID_TETTING_STR = ["Ja", "Nei"]
VALID_FASE = ["Fase 1", "Fase 2"]
VALID_STATUS = ["Godkjent", "Ikke godkjent", "Behandles av RIB"]
VALID_FUNKSJON = ["Bæring", "Vanntetting", "Brann", "Lyd"]

ID_PATTERN = re.compile(r'^[A-Z]{2,4}_Hull-\d+$')
DATE_PATTERN = re.compile(r'^\d{2}\.\d{2}\.\d{4}$')
REV_PATTERN = re.compile(r'^\d+$')


@dataclass
class ValidationResult:
    is_valid: bool
    message: str
    severity: str = "error"


@dataclass
class ElementValidation:
    guid: str
    element_name: str
    element_type: str
    object_type: Optional[str]
    has_pset: bool
    pset_name_found: Optional[str]
    all_psets: List[str]
    properties_found: Dict[str, Any]
    property_validations: Dict[str, ValidationResult]
    dimension_validation: Optional[ValidationResult]
    overall_status: str
    error_count: int = 0
    warning_count: int = 0


@dataclass
class FileValidation:
    filename: str
    total_elements: int
    elements: List[ElementValidation]
    summary: Dict[str, int] = field(default_factory=dict)

    def calculate_summary(self):
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


# ============================================================================
# Validation Functions
# ============================================================================

def validate_kategori(value):
    if value is None or str(value).strip() == "":
        return ValidationResult(False, "Kategori mangler", "error")
    if str(value) in VALID_KATEGORI:
        return ValidationResult(True, "OK")
    return ValidationResult(False, f"Ugyldig kategori: '{value}'", "error")


def validate_id(value, expected_prefix=None):
    if value is None or str(value).strip() == "":
        return ValidationResult(False, "ID mangler", "error")
    val_str = str(value).strip()
    if not ID_PATTERN.match(val_str):
        return ValidationResult(False, f"Ugyldig ID-format: '{val_str}'", "error")
    if expected_prefix:
        prefix = val_str.split('_')[0]
        if prefix != expected_prefix:
            return ValidationResult(True, f"ID-prefiks '{prefix}' != '{expected_prefix}'", "warning")
    return ValidationResult(True, "OK")


def validate_utsparingstype(value):
    if value is None or str(value).strip() == "":
        return ValidationResult(False, "Utsparingstype mangler", "error")
    val_str = str(value).strip()
    if val_str in ["Utsparing", "Hulltaking"]:
        return ValidationResult(True, "OK")
    if val_str == "Innstøpningsgods":
        return ValidationResult(True, "Innstøpningsgods - skal normalt ikke være i utsparings-IFC", "warning")
    return ValidationResult(False, f"Ugyldig utsparingstype: '{val_str}'", "error")


def validate_tetting(value):
    if value is None:
        return ValidationResult(False, "Tetting mangler", "error")
    if isinstance(value, bool):
        return ValidationResult(True, "OK")
    val_str = str(value).strip()
    if val_str == "":
        return ValidationResult(False, "Tetting mangler", "error")
    if val_str in VALID_TETTING_STR:
        return ValidationResult(True, "OK")
    return ValidationResult(False, f"Ugyldig tetting-verdi: '{value}'", "error")


def validate_funksjon(value):
    if value is None or str(value).strip() == "":
        return ValidationResult(True, "Funksjon ikke angitt (valgfritt)", "info")
    val_str = str(value).strip()
    values = [v.strip() for v in val_str.split(',')]
    invalid = [v for v in values if v and v not in VALID_FUNKSJON]
    if invalid:
        return ValidationResult(True, f"Ukjent funksjon: {', '.join(invalid)}", "warning")
    return ValidationResult(True, "OK")


def validate_fase(value):
    if value is None or str(value).strip() == "":
        return ValidationResult(False, "Fase mangler", "error")
    if str(value).strip() in VALID_FASE:
        return ValidationResult(True, "OK")
    return ValidationResult(False, f"Ugyldig fase: '{value}'", "error")


def validate_status(value):
    if value is None or str(value).strip() == "":
        return ValidationResult(False, "Status mangler", "error")
    if str(value).strip() in VALID_STATUS:
        return ValidationResult(True, "OK")
    return ValidationResult(False, f"Ugyldig status: '{value}'", "error")


def validate_rev(value):
    if value is None:
        return ValidationResult(False, "Revisjon mangler", "error")
    val_str = str(value).strip()
    if val_str == "":
        return ValidationResult(False, "Revisjon mangler", "error")
    if REV_PATTERN.match(val_str):
        return ValidationResult(True, "OK")
    try:
        if int(float(value)) >= 0:
            return ValidationResult(True, "OK")
    except (ValueError, TypeError):
        pass
    return ValidationResult(False, f"Ugyldig revisjonsnummer: '{value}'", "error")


def validate_rev_dato(value):
    if value is None or str(value).strip() == "":
        return ValidationResult(False, "Revisjonsdato mangler", "error")
    if DATE_PATTERN.match(str(value).strip()):
        return ValidationResult(True, "OK")
    return ValidationResult(False, f"Ugyldig datoformat: '{value}' (forventet: DD.MM.YYYY)", "error")


def validate_rev_beskrivelse(value):
    if value is None or str(value).strip() == "":
        return ValidationResult(False, "Revisjonsbeskrivelse mangler", "error")
    return ValidationResult(True, "OK")


def validate_dimensions(properties):
    bredde = properties.get("A4_Utsp_DimBredde")
    hoyde = properties.get("A4_Utsp_DimHøyde")
    diameter = properties.get("A4_Utsp_DimDiameter")

    def has_value(val):
        if val is None:
            return False
        val_str = str(val).strip()
        return val_str != "" and val_str != "0" and val_str != "0.0"

    if has_value(bredde) and has_value(hoyde):
        return ValidationResult(True, "Rektangulær: Bredde og Høyde angitt")
    if has_value(diameter):
        return ValidationResult(True, "Rund: Diameter angitt")
    if has_value(bredde) and not has_value(hoyde):
        return ValidationResult(False, "Bredde angitt men Høyde mangler", "error")
    if has_value(hoyde) and not has_value(bredde):
        return ValidationResult(False, "Høyde angitt men Bredde mangler", "error")
    return ValidationResult(False, "Dimensjoner mangler (enten Bredde+Høyde eller Diameter)", "error")


PROPERTY_VALIDATORS = {
    "A4_Utsp_Kategori": validate_kategori,
    "A4_Utsp_ID": validate_id,
    "A4_Utsp_Utsparingstype": validate_utsparingstype,
    "A4_Utsp_Tetting": validate_tetting,
    "A4_Utsp_Funksjon": validate_funksjon,
    "A4_Utsp_Fase": validate_fase,
    "A4_Utsp_Status": validate_status,
    "A4_Utsp_Rev": validate_rev,
    "A4_Utsp_RevDato": validate_rev_dato,
    "A4_Utsp_RevBeskrivelse": validate_rev_beskrivelse,
}


def extract_file_prefix(filename: str) -> Optional[str]:
    name = Path(filename).stem
    parts = name.split('_')
    for part in parts:
        if part in ['RIV', 'RIE', 'RIVA']:
            return part
    return None


def get_element_properties(element) -> Tuple[Optional[str], Dict[str, Any], List[str]]:
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

        if pset_name == REQUIRED_PSET:
            pset_name_found = REQUIRED_PSET
        elif any(prop.Name.startswith("A4_Utsp_") for prop in prop_def.HasProperties if hasattr(prop, 'Name')):
            if pset_name_found is None:
                pset_name_found = f"{pset_name} (feil plassering)"

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
    guid = element.GlobalId
    element_name = element.Name or "(uten navn)"
    element_type = element.is_a()
    object_type = getattr(element, 'ObjectType', None)

    pset_name_found, properties, all_psets = get_element_properties(element)
    has_pset = pset_name_found == REQUIRED_PSET

    property_validations = {}
    error_count = 0
    warning_count = 0

    if pset_name_found is None:
        return ElementValidation(
            guid=guid, element_name=element_name, element_type=element_type,
            object_type=object_type, has_pset=False, pset_name_found=None,
            all_psets=all_psets, properties_found={}, property_validations={},
            dimension_validation=None, overall_status="Mangler pset",
            error_count=1, warning_count=0
        )

    for prop_name in REQUIRED_PROPERTIES:
        value = properties.get(prop_name)
        if prop_name in PROPERTY_VALIDATORS:
            validator = PROPERTY_VALIDATORS[prop_name]
            if prop_name == "A4_Utsp_ID":
                result = validate_id(value, expected_prefix)
            else:
                result = validator(value)
        else:
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

    for prop_name in OPTIONAL_PROPERTIES:
        if prop_name in properties:
            validator = PROPERTY_VALIDATORS.get(prop_name)
            if validator:
                result = validator(properties[prop_name])
                property_validations[prop_name] = result
                if result.severity == "warning":
                    warning_count += 1

    dim_result = validate_dimensions(properties)
    if not dim_result.is_valid:
        error_count += 1

    if pset_name_found and pset_name_found != REQUIRED_PSET:
        warning_count += 1

    if error_count > 0:
        overall_status = "Feil"
    elif warning_count > 0:
        overall_status = "Advarsel"
    else:
        overall_status = "OK"

    return ElementValidation(
        guid=guid, element_name=element_name, element_type=element_type,
        object_type=object_type, has_pset=has_pset, pset_name_found=pset_name_found,
        all_psets=all_psets, properties_found=properties,
        property_validations=property_validations, dimension_validation=dim_result,
        overall_status=overall_status, error_count=error_count, warning_count=warning_count
    )


def validate_ifc_file(file_content: bytes, filename: str) -> FileValidation:
    with tempfile.NamedTemporaryFile(suffix='.ifc', delete=False) as tmp:
        tmp.write(file_content)
        tmp_path = tmp.name

    model = ifcopenshell.open(tmp_path)
    expected_prefix = extract_file_prefix(filename)
    elements = list(model.by_type("IfcBuildingElementProxy"))

    validated_elements = []
    for element in elements:
        validation = validate_element(element, expected_prefix)
        validated_elements.append(validation)

    file_validation = FileValidation(
        filename=filename,
        total_elements=len(elements),
        elements=validated_elements
    )
    file_validation.calculate_summary()

    Path(tmp_path).unlink()
    return file_validation


def get_error_messages(ev: ElementValidation) -> str:
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

    if not ev.has_pset and ev.pset_name_found is None:
        if ev.all_psets:
            error_messages.append(f"Mangler A4_Utsp. Har: {', '.join(ev.all_psets)}")
        else:
            error_messages.append(f"Mangler A4_Utsp. Sjekk GUID i IFC")
    elif ev.pset_name_found and ev.pset_name_found != "A4_Utsp":
        warning_messages.append(f"Egenskaper i feil pset: {ev.pset_name_found}")

    all_msgs = []
    all_msgs.extend([f"FEIL: {m}" for m in error_messages])
    all_msgs.extend([f"ADVARSEL: {m}" for m in warning_messages])
    return "; ".join(all_msgs) if all_msgs else "-"


def create_validated_ifc(file_content: bytes, filename: str, validation: FileValidation) -> bytes:
    """Create IFC with NOSKI_Validering pset and traffic light colors."""
    with tempfile.NamedTemporaryFile(suffix='.ifc', delete=False) as tmp:
        tmp.write(file_content)
        tmp_path = tmp.name

    model = ifcopenshell.open(tmp_path)

    # Create lookup for validation results
    validation_lookup = {ev.guid: ev for ev in validation.elements}

    # Traffic light colors (RGB 0-1)
    COLORS = {
        "OK": (0.22, 0.80, 0.44),           # Green
        "Advarsel": (0.96, 0.62, 0.04),     # Orange
        "Feil": (0.92, 0.26, 0.20),         # Red
        "Mangler pset": (0.49, 0.27, 0.85), # Purple
    }

    # Create color styles
    styles = {}
    for status, (r, g, b) in COLORS.items():
        colour = model.create_entity("IfcColourRgb", Name=f"Validering_{status}", Red=r, Green=g, Blue=b)
        rendering = model.create_entity("IfcSurfaceStyleRendering",
            SurfaceColour=colour,
            Transparency=0.0,
            ReflectanceMethod="NOTDEFINED"
        )
        surface_style = model.create_entity("IfcSurfaceStyle",
            Name=f"Validering_{status}_Style",
            Side="BOTH",
            Styles=[rendering]
        )
        styles[status] = surface_style

    for element in model.by_type("IfcBuildingElementProxy"):
        if not hasattr(element, 'GlobalId'):
            continue

        guid = element.GlobalId
        if guid not in validation_lookup:
            continue

        ev = validation_lookup[guid]

        # Determine status
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

        dim_status = ev.dimension_validation.message if ev.dimension_validation else "Ikke validert"
        error_summary = get_error_messages(ev)

        # Create NOSKI_Validering pset
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

        # Add individual property checks
        for prop_name, result in ev.property_validations.items():
            short_name = prop_name.replace("A4_Utsp_", "")
            property_values.append(
                model.create_entity("IfcPropertySingleValue",
                    Name=f"{short_name}_sjekk",
                    NominalValue=model.create_entity("IfcLabel", result.message))
            )

        property_set = model.create_entity("IfcPropertySet",
            GlobalId=ifcopenshell.guid.new(),
            Name="NOSKI_Validering",
            HasProperties=property_values)

        model.create_entity("IfcRelDefinesByProperties",
            GlobalId=ifcopenshell.guid.new(),
            RelatedObjects=[element],
            RelatingPropertyDefinition=property_set)

        # Apply traffic light color
        status_style = styles.get(ev.overall_status)
        if status_style and hasattr(element, 'Representation') and element.Representation:
            for rep in element.Representation.Representations:
                if hasattr(rep, 'Items'):
                    for item in rep.Items:
                        # Create styled item
                        model.create_entity("IfcStyledItem",
                            Item=item,
                            Styles=[model.create_entity("IfcPresentationStyleAssignment", Styles=[status_style])]
                        )

    # Write to bytes
    output_path = tmp_path.replace('.ifc', '_validert.ifc')
    model.write(output_path)

    with open(output_path, 'rb') as f:
        result = f.read()

    # Cleanup
    Path(tmp_path).unlink()
    Path(output_path).unlink()

    return result


def create_excel_report(validations: List[FileValidation]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Build summary data
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
                "Dekningsgrad A4_Utsp": f"{s['has_pset'] / s['total'] * 100:.1f}%" if s['total'] > 0 else "0%",
            })

        # Build detail data
        all_rows = []
        for fv in validations:
            for ev in fv.elements:
                all_rows.append({
                    "Fil": fv.filename,
                    "GUID": ev.guid,
                    "Navn": ev.element_name,
                    "ObjectType": ev.object_type or "-",
                    "Status": ev.overall_status,
                    "Feilmeldinger": get_error_messages(ev),
                    "Tilgjengelige psets": ", ".join(ev.all_psets) if ev.all_psets else "-",
                    **{k: ev.properties_found.get(k, "") for k in REQUIRED_PROPERTIES + DIMENSION_PROPERTIES},
                })

        # Write summary at top, then gap, then details
        summary_df = pd.DataFrame(summary_data)
        details_df = pd.DataFrame(all_rows)

        summary_df.to_excel(writer, sheet_name="Rapport", index=False, startrow=0)

        # Start details after summary + 2 empty rows
        detail_start_row = len(summary_df) + 3
        details_df.to_excel(writer, sheet_name="Rapport", index=False, startrow=detail_start_row)

        # Auto-fit column widths
        worksheet = writer.sheets["Rapport"]
        for column_cells in worksheet.columns:
            max_length = 0
            column_letter = column_cells[0].column_letter
            for cell in column_cells:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            adjusted_width = min(max_length + 2, 50)  # Cap at 50 chars
            worksheet.column_dimensions[column_letter].width = adjusted_width

    return output.getvalue()


def create_html_report(validations: List[FileValidation]) -> str:
    """Generate standalone HTML report."""
    js_data = {
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "files": []
    }

    for fv in validations:
        file_data = {"filename": fv.filename, "summary": fv.summary, "elements": []}
        for ev in fv.elements:
            elem_data = {
                "guid": ev.guid,
                "name": ev.element_name,
                "objectType": ev.object_type or "-",
                "status": ev.overall_status,
                "errors": get_error_messages(ev),
                "properties": {k: str(v) if v else "" for k, v in ev.properties_found.items()},
            }
            file_data["elements"].append(elem_data)
        js_data["files"].append(file_data)

    return f'''<!DOCTYPE html>
<html lang="no">
<head>
    <meta charset="UTF-8">
    <title>Skiplum pset-sjekk - Utsparingsmodell</title>
    <style>
        * {{ box-sizing: border-box; margin: 0; padding: 0; }}
        body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: linear-gradient(135deg, #f5f5f0 0%, #e8e4dc 100%); color: #333; min-height: 100vh; }}
        .container {{ max-width: 1000px; margin: 0 auto; padding: 20px; }}
        header {{ background: linear-gradient(135deg, #2d4a3e 0%, #3d5a4e 100%); color: white; padding: 25px; border-radius: 12px; margin-bottom: 20px; }}
        header h1 {{ font-size: 1.5rem; margin-bottom: 5px; }}
        header p {{ opacity: 0.85; font-size: 0.9rem; }}
        .summary-grid {{ display: grid; grid-template-columns: repeat(6, 1fr); gap: 10px; margin-bottom: 20px; }}
        .summary-card {{ background: white; border-radius: 10px; padding: 12px; text-align: center; border-left: 3px solid #cbd5e1; box-shadow: 0 2px 6px rgba(0,0,0,0.06); }}
        .summary-card h3 {{ font-size: 0.7rem; color: #666; text-transform: uppercase; margin-bottom: 4px; }}
        .summary-card .value {{ font-size: 1.4rem; font-weight: 700; }}
        .summary-card.ok {{ border-left-color: #10b981; }} .summary-card.ok .value {{ color: #059669; }}
        .summary-card.error {{ border-left-color: #ef4444; }} .summary-card.error .value {{ color: #dc2626; }}
        .summary-card.warning {{ border-left-color: #f59e0b; }} .summary-card.warning .value {{ color: #d97706; }}
        .summary-card.missing {{ border-left-color: #8b5cf6; }} .summary-card.missing .value {{ color: #7c3aed; }}
        .pset-tile {{ border-radius: 10px; padding: 16px; text-align: center; box-shadow: 0 2px 6px rgba(0,0,0,0.1); }}
        .pset-tile h4 {{ font-size: 0.75rem; text-transform: uppercase; margin-bottom: 6px; opacity: 0.9; }}
        .pset-tile .value {{ font-size: 1.4rem; font-weight: 700; }}
        .pset-tile.ok {{ background: #10b981; color: white; }}
        .pset-tile.warning {{ background: #f59e0b; color: black; }}
        .pset-tile.error {{ background: #ef4444; color: white; }}
        .file-section {{ background: white; border-radius: 12px; padding: 20px; margin-bottom: 15px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }}
        .file-header {{ display: flex; justify-content: space-between; align-items: center; cursor: pointer; padding-bottom: 15px; border-bottom: 2px solid #e2e8f0; margin-bottom: 15px; }}
        .file-header h2 {{ font-size: 1.1rem; color: #2d4a3e; }}
        .file-stats {{ display: flex; gap: 8px; }}
        .file-stat {{ padding: 4px 10px; border-radius: 15px; font-size: 0.8rem; font-weight: 600; }}
        .file-stat.ok {{ background: #dcfce7; color: #166534; }}
        .file-stat.error {{ background: #fee2e2; color: #991b1b; }}
        .file-stat.warning {{ background: #fef3c7; color: #92400e; }}
        .file-stat.missing {{ background: #ede9fe; color: #5b21b6; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 0.85rem; }}
        th {{ background: #f8f9fa; padding: 10px; text-align: left; font-weight: 600; color: #4a5568; border-bottom: 2px solid #e2e8f0; }}
        td {{ padding: 8px 10px; border-bottom: 1px solid #eee; }}
        tr:hover {{ background: #f8f9fa; }}
        .status {{ display: inline-block; padding: 3px 8px; border-radius: 10px; font-size: 0.75rem; font-weight: 600; }}
        .status.ok {{ background: #dcfce7; color: #166534; }}
        .status.feil {{ background: #fee2e2; color: #991b1b; }}
        .status.advarsel {{ background: #fef3c7; color: #92400e; }}
        .status.mangler-pset {{ background: #ede9fe; color: #5b21b6; }}
        .guid {{ font-family: monospace; font-size: 0.8rem; color: #666; }}
        .errors {{ font-size: 0.8rem; color: #991b1b; max-width: 250px; }}
        .filters {{ display: flex; gap: 8px; margin-bottom: 15px; }}
        .filter-btn {{ padding: 6px 12px; border: 2px solid #e2e8f0; border-radius: 6px; background: white; cursor: pointer; font-size: 0.85rem; }}
        .filter-btn:hover {{ border-color: #2d4a3e; }}
        .filter-btn.active {{ background: #2d4a3e; color: white; border-color: #2d4a3e; }}
        .footer {{ text-align: center; padding: 20px; color: #666; font-size: 0.8rem; }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>🏗️ Skiplum pset-sjekk</h1>
            <p>Utsparingsmodell — Generert: {js_data["generated"]}</p>
        </header>
        <div id="summary"></div>
        <div id="files"></div>
        <div class="footer">Valideringsrapport — A4_Utsp</div>
    </div>
    <script>
        const data = {json.dumps(js_data, ensure_ascii=False)};
        const totals = data.files.reduce((a,f) => {{
            a.total += f.summary.total; a.ok += f.summary.ok; a.feil += f.summary.feil;
            a.advarsel += f.summary.advarsel; a.mangler_pset += f.summary.mangler_pset;
            return a;
        }}, {{total:0, ok:0, feil:0, advarsel:0, mangler_pset:0}});
        const pct = totals.total > 0 ? ((totals.ok/totals.total)*100).toFixed(1) : 0;

        document.getElementById('summary').innerHTML = `
            <div class="summary-grid">
                <div class="summary-card"><h3>Totalt</h3><div class="value">${{totals.total}}</div></div>
                <div class="summary-card ok"><h3>OK</h3><div class="value">${{totals.ok}}</div></div>
                <div class="summary-card error"><h3>Feil</h3><div class="value">${{totals.feil}}</div></div>
                <div class="summary-card warning"><h3>Advarsler</h3><div class="value">${{totals.advarsel}}</div></div>
                <div class="summary-card missing"><h3>Mangler</h3><div class="value">${{totals.mangler_pset}}</div></div>
                <div class="summary-card"><h3>Godkjent</h3><div class="value">${{pct}}%</div></div>
            </div>`;

        let html = '';
        data.files.forEach((file, i) => {{
            html += `<div class="file-section">
                <div class="file-header">
                    <h2>${{file.filename}}</h2>
                    <div class="file-stats">
                        <span class="file-stat ok">${{file.summary.ok}} OK</span>
                        <span class="file-stat error">${{file.summary.feil}} Feil</span>
                        <span class="file-stat warning">${{file.summary.advarsel}} Adv</span>
                        <span class="file-stat missing">${{file.summary.mangler_pset}} Mangler</span>
                    </div>
                </div>
                <div class="filters">
                    <button class="filter-btn active" onclick="filter(${{i}},'all',this)">Alle</button>
                    <button class="filter-btn" onclick="filter(${{i}},'ok',this)">OK</button>
                    <button class="filter-btn" onclick="filter(${{i}},'feil',this)">Feil</button>
                    <button class="filter-btn" onclick="filter(${{i}},'advarsel',this)">Advarsler</button>
                    <button class="filter-btn" onclick="filter(${{i}},'mangler-pset',this)">Mangler</button>
                </div>
                <table id="t${{i}}"><thead><tr><th>Status</th><th>ID</th><th>ObjectType</th><th>Feilmeldinger</th><th>GUID</th></tr></thead>
                <tbody>${{file.elements.map(e => `<tr data-s="${{e.status.toLowerCase().replace(' ','-')}}">
                    <td><span class="status ${{e.status.toLowerCase().replace(' ','-')}}">${{e.status}}</span></td>
                    <td>${{e.properties['A4_Utsp_ID'] || '-'}}</td>
                    <td>${{e.objectType}}</td>
                    <td class="errors">${{e.errors || '-'}}</td>
                    <td class="guid">${{e.guid}}</td>
                </tr>`).join('')}}</tbody></table></div>`;
        }});
        document.getElementById('files').innerHTML = html;

        function filter(i, s, btn) {{
            btn.parentElement.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            document.querySelectorAll(`#t${{i}} tbody tr`).forEach(r => {{
                r.style.display = (s === 'all' || r.dataset.s === s) ? '' : 'none';
            }});
        }}
    </script>
</body>
</html>'''


# ============================================================================
# Streamlit App
# ============================================================================

st.set_page_config(
    page_title="Skiplum pset-sjekk",
    page_icon="🏗️",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Custom CSS - warm, professional palette
st.markdown("""
<style>
    /* Warm sage/olive background */
    .stApp { background: linear-gradient(135deg, #f5f5f0 0%, #e8e4dc 100%); }
    .main > div { max-width: 900px; margin: 0 auto; }
    .block-container { padding-top: 4rem; padding-bottom: 2rem; max-width: 900px; }

    /* Header */
    .app-header {
        background: linear-gradient(135deg, #2d4a3e 0%, #3d5a4e 100%);
        color: white;
        padding: 1.5rem 2rem;
        border-radius: 12px;
        margin-bottom: 1rem;
    }
    .app-header h1 { color: white; margin: 0; font-size: 1.5rem; font-weight: 600; }
    .app-header p { color: #b8c9bf; margin: 0.25rem 0 0 0; font-size: 0.9rem; }

    /* Summary cards - full width grid */
    .summary-grid {
        display: grid;
        grid-template-columns: repeat(6, 1fr);
        gap: 0.75rem;
        margin-bottom: 1rem;
    }
    .summary-card {
        background: white;
        border-radius: 10px;
        padding: 0.75rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        text-align: center;
        border-left: 3px solid #cbd5e1;
    }
    .summary-card h4 { font-size: 0.7rem; color: #64748b; margin: 0 0 0.25rem 0; text-transform: uppercase; }
    .summary-card .value { font-size: 1.5rem; font-weight: 700; color: #334155; }
    .summary-card.ok { border-left-color: #10b981; }
    .summary-card.ok .value { color: #059669; }
    .summary-card.error { border-left-color: #ef4444; }
    .summary-card.error .value { color: #dc2626; }
    .summary-card.warning { border-left-color: #f59e0b; }
    .summary-card.warning .value { color: #d97706; }
    .summary-card.missing { border-left-color: #8b5cf6; }
    .summary-card.missing .value { color: #7c3aed; }

    /* Pset check tile - solid background */
    .pset-tile {
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    .pset-tile h4 {
        font-size: 0.7rem;
        text-transform: uppercase;
        margin: 0 0 0.5rem 0;
        opacity: 0.9;
    }
    .pset-tile .value {
        font-size: 1.4rem;
        font-weight: 700;
    }
    .pset-tile.ok { background: #10b981; color: white; }
    .pset-tile.warning { background: #f59e0b; color: black; }
    .pset-tile.error { background: #ef4444; color: white; }

    /* Streamlit overrides */
    .stAlert { border-radius: 10px; }
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }

    /* File uploader styling */
    [data-testid="stFileUploader"] {
        background: white;
        padding: 1rem;
        border-radius: 10px;
        border: 2px dashed #94a3b8;
    }

    /* Tabs styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: transparent;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 8px 16px;
        border-radius: 8px;
        border: 2px solid #cbd5e1;
        background-color: white;
        color: #475569;
    }
    .stTabs [data-baseweb="tab"]:hover {
        border-color: #2563eb;
        color: #2563eb;
    }
    .stTabs [aria-selected="true"] {
        background-color: #2563eb !important;
        color: white !important;
        border-color: #2563eb !important;
    }
    .stTabs [data-baseweb="tab-highlight"] {
        display: none;
    }
    .stTabs [data-baseweb="tab-border"] {
        display: none;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
<div class="app-header">
    <h1>🏗️ Skiplum pset-sjekk</h1>
    <p>Utsparingsmodell — A4_Utsp validering</p>
</div>
""", unsafe_allow_html=True)

# File upload - full width initially, split after upload
uploaded_files = st.file_uploader(
    "Last opp IFC-filer med utsparinger",
    type=['ifc'],
    accept_multiple_files=True
)

if uploaded_files:
    # Validate files and store content for IFC export
    with st.spinner("Validerer IFC-filer..."):
        validations = []
        file_contents = {}
        for uploaded_file in uploaded_files:
            content = uploaded_file.read()
            file_contents[uploaded_file.name] = content
            validation = validate_ifc_file(content, uploaded_file.name)
            validations.append(validation)

    # Check pset status per file - one tile each (up to 4)
    def get_file_pset_status(fv):
        has_correct = any(e.has_pset for e in fv.elements)
        has_wrong = any(e.pset_name_found and "(feil plassering)" in e.pset_name_found for e in fv.elements)
        if has_correct:
            return "ok", "A4_Utsp", None
        elif has_wrong:
            for e in fv.elements:
                if e.pset_name_found and "(feil plassering)" in e.pset_name_found:
                    return "warning", e.pset_name_found.replace(" (feil plassering)", ""), "Feil pset-navn"
            return "warning", "Ukjent", "Feil pset-navn"
        else:
            return "error", "Ikke funnet", None

    # Build tiles HTML in a flex row
    tile_parts = []
    for fv in validations[:4]:  # Max 4 files
        status_class, status_text, subtitle = get_file_pset_status(fv)
        filename = fv.filename.replace(".ifc", "")
        subtitle_html = f'<div style="font-size: 0.75rem; margin-top: 4px; opacity: 0.85;">{subtitle}</div>' if subtitle else ''
        tile_parts.append(f'<div class="pset-tile {status_class}" style="flex: 1;"><h4>{filename}</h4><div class="value">{status_text}</div>{subtitle_html}</div>')

    tiles_html = f'<div style="display: flex; gap: 0.75rem; margin-bottom: 1rem;">{"".join(tile_parts)}</div>'
    st.markdown(tiles_html, unsafe_allow_html=True)

    # Overall summary using custom HTML cards
    totals = {
        "total": sum(v.summary["total"] for v in validations),
        "ok": sum(v.summary["ok"] for v in validations),
        "feil": sum(v.summary["feil"] for v in validations),
        "advarsel": sum(v.summary["advarsel"] for v in validations),
        "mangler_pset": sum(v.summary["mangler_pset"] for v in validations),
    }
    pct = (totals["ok"] / totals["total"] * 100) if totals["total"] > 0 else 0

    st.markdown(f"""
    <div class="summary-grid">
        <div class="summary-card">
            <h4>Totalt</h4>
            <div class="value">{totals["total"]}</div>
        </div>
        <div class="summary-card ok">
            <h4>OK</h4>
            <div class="value">{totals["ok"]}</div>
        </div>
        <div class="summary-card error">
            <h4>Feil</h4>
            <div class="value">{totals["feil"]}</div>
        </div>
        <div class="summary-card warning">
            <h4>Advarsler</h4>
            <div class="value">{totals["advarsel"]}</div>
        </div>
        <div class="summary-card missing">
            <h4>Mangler pset</h4>
            <div class="value">{totals["mangler_pset"]}</div>
        </div>
        <div class="summary-card">
            <h4>Godkjenningsgrad</h4>
            <div class="value">{pct:.1f}%</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Downloads row - 3 columns
    excel_data = create_excel_report(validations)
    html_data = create_html_report(validations)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")

    dl_cols = st.columns(3)
    with dl_cols[0]:
        st.download_button(
            label="📊 Last ned Excel",
            data=excel_data,
            file_name=f"A4_Utsp_validering_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with dl_cols[1]:
        st.download_button(
            label="📄 Last ned HTML",
            data=html_data,
            file_name=f"A4_Utsp_validering_{timestamp}.html",
            mime="text/html",
            use_container_width=True
        )
    with dl_cols[2]:
        # IFC download - single or multiple files
        if len(validations) == 1:
            fv = validations[0]
            ifc_data = create_validated_ifc(file_contents[fv.filename], fv.filename, fv)
            st.download_button(
                label="🏗️ Last ned IFC",
                data=ifc_data,
                file_name=f"{Path(fv.filename).stem}_validert.ifc",
                mime="application/octet-stream",
                use_container_width=True
            )
        else:
            # Multiple files - popover with individual downloads
            with st.popover("🏗️ Last ned IFC", use_container_width=True):
                for idx, fv in enumerate(validations):
                    ifc_data = create_validated_ifc(file_contents[fv.filename], fv.filename, fv)
                    stem = Path(fv.filename).stem
                    st.download_button(
                        label=f"{stem}_validert.ifc",
                        data=ifc_data,
                        file_name=f"{stem}_validert.ifc",
                        mime="application/octet-stream",
                        key=f"ifc_dl_{idx}",
                        use_container_width=True
                    )

    # Store validations in session state for fragment access
    st.session_state['validations'] = validations

    # Full-screen dialog for viewing results
    @st.dialog("Valideringsresultater", width="large")
    def show_results_dialog():
        vals = st.session_state.get('validations', [])

        # Compact summary at top of dialog
        totals = st.session_state.get('totals', {})
        st.markdown(f"**{totals.get('total', 0)}** elementer: "
                   f"✅ {totals.get('ok', 0)} OK · "
                   f"❌ {totals.get('feil', 0)} Feil · "
                   f"⚠️ {totals.get('advarsel', 0)} Advarsler · "
                   f"🟣 {totals.get('mangler_pset', 0)} Mangler")

        filter_tabs = st.tabs(["Alle", "OK", "Feil", "Advarsel", "Mangler pset"])

        for tab, status_filter in zip(filter_tabs, ["Alle", "OK", "Feil", "Advarsel", "Mangler pset"]):
            with tab:
                rows = []
                for fv in vals:
                    for ev in fv.elements:
                        if status_filter != "Alle" and ev.overall_status != status_filter:
                            continue
                        pset_display = ev.pset_name_found or "-"
                        if pset_display and "(feil plassering)" in pset_display:
                            pset_display = pset_display.replace(" (feil plassering)", " ⚠️")
                        rows.append({
                            "Fil": fv.filename.replace(".ifc", ""),
                            "Pset": pset_display,
                            "Status": ev.overall_status,
                            "A4_Utsp_ID": ev.properties_found.get("A4_Utsp_ID", "-"),
                            "ObjectType": ev.object_type or "-",
                            "Feil": get_error_messages(ev) or "-",
                            "GUID": ev.guid,
                            # Extra data for detail view (JSON strings for Arrow compatibility)
                            "_all_psets": json.dumps(ev.all_psets),
                            "_properties": json.dumps({k: str(v) if v is not None else None for k, v in ev.properties_found.items()}),
                            "_validations": json.dumps({k: [v.is_valid, v.message, v.severity] for k, v in ev.property_validations.items()}),
                        })

                if rows:
                    df = pd.DataFrame(rows)

                    def color_status(val):
                        colors = {
                            "OK": "background-color: #dcfce7; color: #166534",
                            "Feil": "background-color: #fee2e2; color: #991b1b",
                            "Advarsel": "background-color: #fef3c7; color: #92400e",
                            "Mangler pset": "background-color: #ede9fe; color: #5b21b6",
                        }
                        return colors.get(val, "")

                    styled_df = df.style.map(color_status, subset=["Status"])

                    event = st.dataframe(
                        styled_df,
                        hide_index=True,
                        height=500,
                        on_select="rerun",
                        selection_mode="single-row",
                        column_config={
                            "Fil": st.column_config.TextColumn("Fil", width="small"),
                            "Pset": st.column_config.TextColumn("Pset", width="small"),
                            "Feil": st.column_config.TextColumn("Feil", width="large"),
                            "GUID": st.column_config.TextColumn("GUID", width="small"),
                            # Hide internal columns
                            "_all_psets": None,
                            "_properties": None,
                            "_validations": None,
                        },
                        key=f"dialog_df_{status_filter}"
                    )

                    if event.selection and event.selection.rows:
                        selected_idx = event.selection.rows[0]
                        selected_row = df.iloc[selected_idx]
                        with st.container(border=True):
                            st.markdown(f"**{selected_row['A4_Utsp_ID']}** — {selected_row['Status']}")
                            st.caption(f"GUID: `{selected_row['GUID']}`")

                            if selected_row['Feil'] != "-":
                                st.error(selected_row['Feil'])

                            # Show detected pset info
                            col1, col2 = st.columns(2)
                            with col1:
                                st.markdown("**Detektert Pset:**")
                                pset_name = selected_row['Pset']
                                if pset_name == "-":
                                    st.warning("Ingen A4_Utsp pset funnet")
                                elif "⚠️" in pset_name:
                                    st.warning(f"`{pset_name}` (feil plassering)")
                                else:
                                    st.success(f"`{pset_name}`")

                            with col2:
                                st.markdown("**Alle psets på element:**")
                                all_psets = json.loads(selected_row['_all_psets'])
                                if all_psets:
                                    st.text("\n".join(all_psets))
                                else:
                                    st.text("(ingen)")

                            # Show properties with validation status
                            st.markdown("**Egenskaper:**")
                            props = json.loads(selected_row['_properties'])
                            validations = json.loads(selected_row['_validations'])

                            if props:
                                prop_rows = []
                                for prop_name, value in props.items():
                                    val_info = validations.get(prop_name, [True, "OK", "ok"])
                                    is_valid, message, severity = val_info
                                    status_icon = "✅" if is_valid and severity != "warning" else "⚠️" if severity == "warning" else "❌"
                                    prop_rows.append({
                                        "": status_icon,
                                        "Egenskap": prop_name,
                                        "Verdi": str(value) if value is not None else "(tom)",
                                        "Validering": message
                                    })
                                st.dataframe(
                                    pd.DataFrame(prop_rows),
                                    hide_index=True,
                                    key=f"props_{status_filter}_{selected_idx}"
                                )
                            else:
                                st.info("Ingen A4_Utsp egenskaper funnet")
                else:
                    st.info("Ingen elementer i denne kategorien.")

    # Store totals for dialog
    st.session_state['totals'] = totals

    # Red button to open dialog
    st.markdown("""
    <style>
    div.stButton > button[kind="primary"] {
        background-color: #dc2626;
        border-color: #dc2626;
    }
    div.stButton > button[kind="primary"]:hover {
        background-color: #b91c1c;
        border-color: #b91c1c;
    }
    </style>
    """, unsafe_allow_html=True)
    if st.button("📋 Se resultater i fullskjerm", use_container_width=True, type="primary"):
        show_results_dialog()

else:
    st.info("Last opp IFC-filer med utsparinger for å validere mot A4_Utsp krav.")

    with st.expander("ℹ️ Valideringskrav"):
        st.markdown("""
**Påkrevd pset:** `A4_Utsp` · **Egenskaper:** Kategori, ID, Utsparingstype, Tetting, Fase, Status, Rev, RevDato, RevBeskrivelse · **Dimensjoner:** Bredde+Høyde eller Diameter
        """)
