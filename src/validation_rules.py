# -*- coding: utf-8 -*-
"""
Validation rules for A4_Utsp property set validation.

Based on BIM-Gjennomføringsplan requirements for utsparinger/hulltaking.
"""

import re
from typing import Dict, Any, Optional, Tuple

# Required property set name
REQUIRED_PSET = "A4_Utsp"

# Required properties in A4_Utsp
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

# Dimension properties (conditional requirement)
DIMENSION_PROPERTIES = [
    "A4_Utsp_DimBredde",
    "A4_Utsp_DimHøyde",
    "A4_Utsp_DimDybde",
    "A4_Utsp_DimDiameter",
]

# Optional properties
OPTIONAL_PROPERTIES = [
    "A4_Utsp_Funksjon",
]

# All properties
ALL_PROPERTIES = REQUIRED_PROPERTIES + DIMENSION_PROPERTIES + OPTIONAL_PROPERTIES

# Valid values for enum fields
VALID_KATEGORI = ["ProvisionForVoid"]
VALID_UTSPARINGSTYPE = ["Utsparing", "Hulltaking", "Innstøpningsgods"]
VALID_UTSPARINGSTYPE_PREFERRED = ["Utsparing", "Hulltaking"]  # Innstøpningsgods triggers warning
VALID_TETTING_BOOL = [True, False]
VALID_TETTING_STR = ["Ja", "Nei"]
VALID_FASE = ["Fase 1", "Fase 2"]
VALID_STATUS = ["Godkjent", "Ikke godkjent", "Behandles av RIB"]
VALID_FUNKSJON = ["Bæring", "Vanntetting", "Brann", "Lyd"]

# ID pattern: {FAG}_Hull-{nn} where FAG is 2-4 uppercase letters
ID_PATTERN = re.compile(r'^[A-Z]{2,4}_Hull-\d+$')

# Date pattern: DD.MM.YYYY
DATE_PATTERN = re.compile(r'^\d{2}\.\d{2}\.\d{4}$')

# Revision pattern: integer 0, 1, 2, ...
REV_PATTERN = re.compile(r'^\d+$')


class ValidationResult:
    """Result of validating a single property or element."""

    def __init__(self, is_valid: bool, message: str, severity: str = "error"):
        """
        Args:
            is_valid: Whether validation passed
            message: Description of result (Norwegian)
            severity: "error", "warning", or "info"
        """
        self.is_valid = is_valid
        self.message = message
        self.severity = severity

    def __repr__(self):
        status = "OK" if self.is_valid else self.severity.upper()
        return f"[{status}] {self.message}"


def validate_kategori(value: Any) -> ValidationResult:
    """Validate A4_Utsp_Kategori."""
    if value is None or str(value).strip() == "":
        return ValidationResult(False, "Kategori mangler", "error")
    if str(value) in VALID_KATEGORI:
        return ValidationResult(True, "OK")
    return ValidationResult(False, f"Ugyldig kategori: '{value}' (forventet: ProvisionForVoid)", "error")


def validate_id(value: Any, expected_prefix: str = None) -> ValidationResult:
    """Validate A4_Utsp_ID format."""
    if value is None or str(value).strip() == "":
        return ValidationResult(False, "ID mangler", "error")

    val_str = str(value).strip()
    if not ID_PATTERN.match(val_str):
        return ValidationResult(False, f"Ugyldig ID-format: '{val_str}' (forventet: XXX_Hull-nn)", "error")

    # Check prefix matches expected (e.g., RIV, RIE, RIVA)
    if expected_prefix:
        prefix = val_str.split('_')[0]
        if prefix != expected_prefix:
            return ValidationResult(False, f"ID-prefiks '{prefix}' matcher ikke forventet '{expected_prefix}'", "warning")

    return ValidationResult(True, "OK")


def validate_utsparingstype(value: Any) -> ValidationResult:
    """Validate A4_Utsp_Utsparingstype."""
    if value is None or str(value).strip() == "":
        return ValidationResult(False, "Utsparingstype mangler", "error")

    val_str = str(value).strip()
    if val_str in VALID_UTSPARINGSTYPE_PREFERRED:
        return ValidationResult(True, "OK")
    if val_str == "Innstøpningsgods":
        return ValidationResult(True, "Innstøpningsgods skal normalt ikke være i utsparings-IFC", "warning")
    return ValidationResult(False, f"Ugyldig utsparingstype: '{val_str}'", "error")


def validate_tetting(value: Any) -> ValidationResult:
    """Validate A4_Utsp_Tetting (accepts both boolean and string)."""
    if value is None:
        return ValidationResult(False, "Tetting mangler", "error")

    # Boolean check
    if isinstance(value, bool):
        return ValidationResult(True, "OK")

    # String check
    val_str = str(value).strip()
    if val_str == "":
        return ValidationResult(False, "Tetting mangler", "error")
    if val_str in VALID_TETTING_STR:
        return ValidationResult(True, "OK")

    return ValidationResult(False, f"Ugyldig tetting-verdi: '{value}' (forventet: Ja/Nei eller True/False)", "error")


def validate_funksjon(value: Any) -> ValidationResult:
    """Validate A4_Utsp_Funksjon (optional, can have multiple values)."""
    if value is None or str(value).strip() == "":
        return ValidationResult(True, "Funksjon ikke angitt (valgfritt)", "info")

    val_str = str(value).strip()
    # Can contain multiple values separated by comma
    values = [v.strip() for v in val_str.split(',')]

    invalid_values = [v for v in values if v and v not in VALID_FUNKSJON]
    if invalid_values:
        return ValidationResult(True, f"Ukjent funksjon: {', '.join(invalid_values)}", "warning")

    return ValidationResult(True, "OK")


def validate_fase(value: Any) -> ValidationResult:
    """Validate A4_Utsp_Fase."""
    if value is None or str(value).strip() == "":
        return ValidationResult(False, "Fase mangler", "error")

    val_str = str(value).strip()
    if val_str in VALID_FASE:
        return ValidationResult(True, "OK")
    return ValidationResult(False, f"Ugyldig fase: '{val_str}' (forventet: Fase 1 eller Fase 2)", "error")


def validate_status(value: Any) -> ValidationResult:
    """Validate A4_Utsp_Status."""
    if value is None or str(value).strip() == "":
        return ValidationResult(False, "Status mangler", "error")

    val_str = str(value).strip()
    if val_str in VALID_STATUS:
        return ValidationResult(True, "OK")
    return ValidationResult(False, f"Ugyldig status: '{val_str}'", "error")


def validate_rev(value: Any) -> ValidationResult:
    """Validate A4_Utsp_Rev (revision number)."""
    if value is None:
        return ValidationResult(False, "Revisjon mangler", "error")

    val_str = str(value).strip()
    if val_str == "":
        return ValidationResult(False, "Revisjon mangler", "error")

    # Accept integer or string representation of integer
    if REV_PATTERN.match(val_str):
        return ValidationResult(True, "OK")

    # Try to parse as number
    try:
        rev_num = int(float(value))
        if rev_num >= 0:
            return ValidationResult(True, "OK")
    except (ValueError, TypeError):
        pass

    return ValidationResult(False, f"Ugyldig revisjonsnummer: '{value}' (forventet: 0, 1, 2, ...)", "error")


def validate_rev_dato(value: Any) -> ValidationResult:
    """Validate A4_Utsp_RevDato (date format DD.MM.YYYY)."""
    if value is None or str(value).strip() == "":
        return ValidationResult(False, "Revisjonsdato mangler", "error")

    val_str = str(value).strip()
    if DATE_PATTERN.match(val_str):
        return ValidationResult(True, "OK")

    return ValidationResult(False, f"Ugyldig datoformat: '{val_str}' (forventet: DD.MM.YYYY)", "error")


def validate_rev_beskrivelse(value: Any) -> ValidationResult:
    """Validate A4_Utsp_RevBeskrivelse (any non-empty text)."""
    if value is None or str(value).strip() == "":
        return ValidationResult(False, "Revisjonsbeskrivelse mangler", "error")
    return ValidationResult(True, "OK")


def validate_dimensions(properties: Dict[str, Any]) -> ValidationResult:
    """
    Validate dimension properties.

    Rules:
    - Rectangular shapes: Bredde AND Høyde required
    - Round shapes: Diameter required
    - Dybde is optional but useful for slits
    """
    bredde = properties.get("A4_Utsp_DimBredde")
    hoyde = properties.get("A4_Utsp_DimHøyde")
    dybde = properties.get("A4_Utsp_DimDybde")
    diameter = properties.get("A4_Utsp_DimDiameter")

    def has_value(val):
        if val is None:
            return False
        val_str = str(val).strip()
        return val_str != "" and val_str != "0" and val_str != "0.0"

    has_bredde = has_value(bredde)
    has_hoyde = has_value(hoyde)
    has_diameter = has_value(diameter)

    # Check for rectangular (bredde + høyde)
    if has_bredde and has_hoyde:
        return ValidationResult(True, "Rektangulær: Bredde og Høyde angitt")

    # Check for round (diameter)
    if has_diameter:
        return ValidationResult(True, "Rund: Diameter angitt")

    # Partial rectangular
    if has_bredde and not has_hoyde:
        return ValidationResult(False, "Bredde angitt men Høyde mangler", "error")
    if has_hoyde and not has_bredde:
        return ValidationResult(False, "Høyde angitt men Bredde mangler", "error")

    # Nothing specified
    return ValidationResult(False, "Dimensjoner mangler (enten Bredde+Høyde eller Diameter)", "error")


# Mapping of property names to validation functions
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
