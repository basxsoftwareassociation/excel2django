from django.apps import apps
from django.core.management.base import BaseCommand
from django.db import transaction
from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter


class Command(BaseCommand):
    help = "Import data from excel into django objects"

    def add_arguments(self, parser):
        parser.add_argument("import_file")

        parser.add_argument(
            "--rows",
            action="append",
            help="""Specify range of rows with <START>:<END>. The index starts at 1 and the range includes the element at index END.
Examples:
--rows 10:20 # all rows from 10 to 20 (including 20)
--rows 2: # all rows except the header row
--rows 2:-1 # All rows except the first and the last one

rows can be specified multiple times, all specified ranges will be imported.
If no arguemnt for rows is given all rows will be imported
""",
        )

        parser.add_argument(
            "--field",
            action="append",
            help="""Define how the value of a field should be extracted from a row. The base syntax is:
    <FIELD-SPECIFICATION>:<VALUE-EXPRESSION>

The syntax for FIELD-SPECIFICATION is:
    <APP_NAME>.<MODEL_NAME>.<FIELD_NAME>

In addition different instances of the same model can be distinguished by adding an arbitrary group-name:
    <APP_NAME>.<MODEL_NAME>.<GROUP_NAME>.<FIELD_NAME>

If the field should be used as component of the natural key the character + should be prefixed:
    +<APP_NAME>.<MODEL_NAME>.<FIELD_NAME>

The FIELD-SPECIFICATION must be unique over all --field arguments except for One-To-Many or Many-To-Many fields.

The value of VALUE-EXPRESSION will be evaluated as python expression. All columns are available as variables with the according volumn name like A, B, AA, etc. Apart from the python builtins there are a few helper functions help with transformations:
    ref(<VALUE>, <APP_NAME>.<MODEL_NAME>.<FIELD1_NAME>.<FIELD2_NAME>) # converts the given VALUE into a reference to <APP_NAME>.<MODEL_NAME> using the <FIELD_NAME>s as natural to look up the object.
    map(<INPUT>, <SRC>, <DST>) # Returns DST if INPUT == SRC otherwise returns INPUT

Examples:
    --field example.Publisher.name:B \\
    --field example.Author.1.first_name:C \\
    --field +example.Author.1.email:D \\
    --field example.Author.2.first_name:E \\
    --field +example.Author.2.email:F \\
    --field example.Book.title:A \\
    --field 'example.Book.publisher:B|ref(example.Publisher.name)'
""",
        )
        parser.add_argument(
            "--sheet",
            help="Specify which sheet to use. Can be a 1-base index or the name of sheet. The default is the first sheet in the excel file.",
        )

    def handle(self, *args, **options):
        rowranges = [(None, None)]
        sheet = try_int(options.get("sheet") or 1)
        if options["rows"]:
            rowranges = combine_ranges(
                (
                    (
                        try_int(r.split(":", 1)[0], None),
                        try_int(r.split(":", 1)[1], None),
                    )
                    for r in options["rows"]
                )
            )

        workbook = load_workbook(options["import_file"], data_only=True, read_only=True)

        if isinstance(sheet, int):
            sheet = workbook.sheetnames[sheet - 1]
        worksheet = workbook.get_sheet_by_name(sheet)

        modeldefinitions = {}
        if "field" in options:
            for fieldarg in options["field"]:
                fieldspec, valueexpr = fieldarg.split(":", 1)
                app_model, field = fieldspec.rsplit(".", 1)
                is_natural_key = app_model.startswith("+")
                app_model = app_model[1:] if is_natural_key else app_model
                app_label, modelname = app_model.split(".", 1)
                if app_model not in modeldefinitions:
                    modeldefinitions[app_model] = {
                        "created": 0,  # for status output
                        "processed": 0,  # for status output
                        "model": apps.get_model(app_label, modelname.split(".", 1)[0]),
                        "fields": {},
                    }
                modeldefinitions[app_model]["fields"][field] = {
                    "modelfield": modeldefinitions[app_model]["model"]._meta.get_field(
                        field
                    ),
                    "expression": valueexpr,
                    "is_natural_key": is_natural_key,
                }
        # if no fields from the command line have been marked as natural key, use all fields as part of the natural key
        for model in modeldefinitions:
            if not any(
                [
                    f["is_natural_key"]
                    for f in modeldefinitions[model]["fields"].values()
                ]
            ):
                for f in modeldefinitions[model]["fields"].values():
                    f["is_natural_key"] = True

        # TODO: create correct order so that referenced objects are created first
        with transaction.atomic():
            for r in rowranges:
                for row in worksheet.iter_rows(min_row=r[0], max_row=r[1]):
                    print(row)
                    rowcontext = {
                        get_column_letter(i): c.value for i, c in enumerate(row, 1)
                    }
                    for model in modeldefinitions:
                        created = importinstance(modeldefinitions[model], rowcontext)
                        modeldefinitions[model]["processed"] += 1
                        if created:
                            modeldefinitions[model]["created"] += 1
        for model in modeldefinitions:
            print(f"{model}:")
            print(f"  {modeldefinitions[model]['created']} created")
            print(f"  {modeldefinitions[model]['processed']} processed")


def importinstance(model, rowcontext):
    naturalkey_values = {
        f["modelfield"].name: _extract_fieldvalue(f, rowcontext)
        for f in model["fields"].values()
        if f["is_natural_key"]
    }
    default_values = {
        f["modelfield"].name: _extract_fieldvalue(f, rowcontext)
        for f in model["fields"].values()
        if not f["is_natural_key"]
    }
    print(default_values, naturalkey_values)
    return model["model"].objects.update_or_create(
        defaults=default_values, **naturalkey_values
    )[1]


def _extract_fieldvalue(field, rowcontext):
    return eval(field["expression"], {}, rowcontext)


def combine_ranges(ranges):
    combined = []
    for r in sorted(ranges):
        if not combined:
            combined.append(r)
        elif not range_overlap(combined[-1][0], combined[-1][1], r[0], r[1]):
            combined.append(r)
        elif r[1] > combined[-1][1]:
            combined[-1][1] = r[1]


def range_overlap(start1, end1, start2, end2):
    def contains(start, end, i):
        if start is None and end is None:
            return True
        if start is None:
            return i <= end
        if end is None:
            return start <= i
        return start <= i <= end

    return (
        contains(start1, end1, start2)
        or contains(start1, end1, end2)
        or contains(start2, end2, start1)
        or contains(start2, end2, end1)
    )


def try_int(v, default=None):
    try:
        return int(v)
    except (TypeError, ValueError):
        return v if default is None else default
