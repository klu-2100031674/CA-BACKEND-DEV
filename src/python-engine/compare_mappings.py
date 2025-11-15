"""
Compare FRCC2 form fields with actual Excel CC2 structure
"""

# From the inspection, these are the ACTUAL INPUT cells in Excel CC2
excel_input_cells = {
    "General Info": ["I3", "I4", "I5", "I6", "I7", "I8"],
    "Finance": ["I10", "I11", "H12", "H13", "H14"],  # Note: H not I for 12-14
    "Years": ["I16", "I17", "I18", "I19", "I20"],
    "Audited": [
        # I22 is FORMULA!
        "I23", "I24", "I25", "I26", "I27", "I28", "I29", "I30", "I31", "I32",
        # I33 is FORMULA!
        "I34", "I35", 
        # I36 is FORMULA!
        "I37",
        # I38, I39 are FORMULAS!
        "I40", "I41",
        # I42 is FORMULA!
        "I43"
    ],
    "Provisional": [
        # I44 - not in Excel!
        "I45", "I46",
        # I47 is FORMULA!
        "I48", "I49", "I50", "I51", "I52", "I53",
        # I54, I55 are FORMULAS!
        "I56",
        # I57 is FORMULA!
        "I58",
        # I59, I60, I61, I62, I63 are FORMULAS!
        "I64", "I65",
        # I66 is FORMULA!
    ],
    "Assumptions": [
        "I67",
        # I68 not used
        # I69 is FORMULA!
        "I70", "I71", "I72", "I73", "I74", "I75", "I76",
        # I77 not used
        "I78", "I79", "I80",
        # I81 not used
        "I82", "I83", "I84",
        # I85 not used
        # I86-I99 are ALL FORMULAS!
    ]
}

# From our FRCC2Form.jsx, these are what we're trying to send
form_fields = {
    "Audited": list(range(22, 44)),  # R22C2 to R43C2
    "Provisional": list(range(44, 67)),  # R44C2 to R66C2
    "Assumptions": list(range(67, 100))  # R67C2 to R99C2
}

print("=" * 100)
print("MISMATCH ANALYSIS: Form Fields vs Excel Input Cells")
print("=" * 100)

print("\n❌ FIELDS THAT SHOULD NOT BE SENT (Have formulas in Excel):")
print("\nAudited Section:")
formula_audited = [22, 33, 36, 38, 39, 42]
for row in formula_audited:
    print(f"  R{row}C2 (I{row}) - Has formula, should be auto-calculated")

print("\nProvisional Section:")
formula_prov = [44, 47, 54, 55, 57, 59, 60, 61, 62, 63, 66]
for row in formula_prov:
    print(f"  R{row}C2 (I{row}) - Has formula or doesn't exist")

print("\nAssumptions Section:")
formula_assum = [68, 69, 77, 81, 85] + list(range(86, 100))
for row in formula_assum:
    print(f"  R{row}C2 (I{row}) - Has formula or not used")

print("\n" + "=" * 100)
print("CORRECT MAPPING SUMMARY")
print("=" * 100)

print("\n✅ AUDITED - Fields that SHOULD be sent:")
should_send_audited = [23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 34, 35, 37, 40, 41, 43]
print(f"  {should_send_audited}")

print("\n✅ PROVISIONAL - Fields that SHOULD be sent:")
should_send_prov = [45, 46, 48, 49, 50, 51, 52, 53, 56, 58, 64, 65]
print(f"  {should_send_prov}")

print("\n✅ ASSUMPTIONS - Fields that SHOULD be sent:")
should_send_assum = [67, 70, 71, 72, 73, 74, 75, 76, 78, 79, 80, 82, 83, 84]
print(f"  {should_send_assum}")

print("\n⚠️  SPECIAL NOTE:")
print("  - I22 (Turnover Audited) has formula =13401963, but this seems wrong!")
print("    It should probably be an input field. Need to check if this is a template error.")
