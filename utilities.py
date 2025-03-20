def clean_parameters(param_string):
    # Remove line breaks, underscores, and extra spaces
    param_string = param_string.replace("_", "").replace("\n", "").strip()
    
    # Split parameters by comma
    params = param_string.split(",")
    cleaned = []
    for param in params:
        name = param.strip().split(" ")[0]
        cleaned.append(name)
    return cleaned


def normalize_code(code_lines):
    normalized_lines = []
    for line in code_lines:
        line = line.replace('\t', '    ')  # Converti tab in spazi
        line = line.rstrip()               # Rimuovi spazi finali
        normalized_lines.append(line)
    # Rimuovi righe vuote multiple consecutive
    cleaned_lines = []
    empty = False
    for line in normalized_lines:
        if line.strip() == "":
            if not empty:
                cleaned_lines.append("")
                empty = True
        else:
            cleaned_lines.append(line)
            empty = False
    return cleaned_lines