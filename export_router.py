from bdo_aca_writer import write_bdo_aca_output
from excel_writer import write_clean_output
from utils import canonical_bank, clean_text


BDO_BANK_NAMES = {
    "BDO",
    "BDO UNIBANK",
    "BDO UNIBANK INC",
    "BANCO DE ORO",
    "BANCO DE ORO UNIBANK",
    "BANCO DE ORO UNIBANK INC",
}


def is_bdo_bank(bank_name):
    normalized = canonical_bank(bank_name)
    normalized = clean_text(normalized)
    return normalized in BDO_BANK_NAMES


def split_rows_by_bank(cleaned_rows):
    bdo_rows = []
    non_bdo_rows = []

    for row in cleaned_rows:
        if is_bdo_bank(row.get("Beneficiary Bank", "")):
            bdo_rows.append(row)
        else:
            non_bdo_rows.append(row)

    return bdo_rows, non_bdo_rows


def split_error_rows_by_bank(error_rows):
    bdo_errors = []
    non_bdo_errors = []

    for row in error_rows:
        if is_bdo_bank(row.get("Beneficiary Bank", "")):
            bdo_errors.append(row)
        else:
            non_bdo_errors.append(row)

    return bdo_errors, non_bdo_errors


def write_routed_outputs(cleaned_rows, error_rows, original_filename):
    bdo_rows, non_bdo_rows = split_rows_by_bank(cleaned_rows)
    bdo_errors, non_bdo_errors = split_error_rows_by_bank(error_rows)

    outputs = {
        "bdo_rows": len(bdo_rows),
        "non_bdo_rows": len(non_bdo_rows),
        "bdo_output_path": None,
        "non_bdo_output_path": None,
        "bdo_errors": bdo_errors,
        "non_bdo_errors": non_bdo_errors,
    }

    if non_bdo_rows:
        outputs["non_bdo_output_path"] = write_clean_output(
            non_bdo_rows,
            non_bdo_errors,
            original_filename,
        )

    if bdo_rows:
        bdo_output_path, bdo_validation_errors = write_bdo_aca_output(
            bdo_rows,
            original_filename,
        )
        outputs["bdo_output_path"] = bdo_output_path
        outputs["bdo_errors"] = bdo_errors + bdo_validation_errors

    return outputs
