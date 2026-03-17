import re


def find_variable_names(struct_expr):
    names = re.findall(r"\{([^}]+)\}", struct_expr)
    if not names:
        return [struct_expr]
    return names


def condition_is_active(struct_expr, truth_map, var_values):
    """
    Evaluate a condition expression and return True if the condition is active, False otherwise.

    Args:
        struct_expr: The "condition" value for the datapoint being evaluated. This contains ids of datapoints to be evaluated and should be found in the condition key of the survey data object being evaluated.
        truth_map: The "condition_criteria" value for the datapoint being evaluated. A dictionary of variable names and their allowed values. This should be populated with the condition_criteria key of the survey data object being evaluated.
        var_values: The "value" for the datapoint being evaluated. A dictionary of variable names and their actual values found in the sopuce survey data items being referenced in the condition expression.

    Returns:
        True if the condition is active, False otherwise.

    Example:
        struct_expr = '(!{var1}|!{var2})&({var1}|{var3})'
        truth_map = {
            "var1": "Collective action|Individual action",
            "var2": "Yes|Partly",
            "var3": "No",
        }
        var_values = {
            "var1": "Individual action",
            "var2": "No",
            "var3": "No",
        }
    """
    # Extract variable names ({var} tokens inside struct_expr)
    struct_vars = re.findall(r"\{([^}]+)\}", struct_expr)

    # Handle simple case where the condition is a single variable, e.g. var1 or {var1}
    if len(struct_vars) == 0 or (
        len(struct_vars) == 1 and struct_vars[0] == struct_expr
    ):
        # Bare name without braces
        var_name = struct_vars[0] if struct_vars else struct_expr
        if isinstance(truth_map, str):
            truth_map = {var_name: truth_map}
        if isinstance(var_values, str):
            var_values = {var_name: var_values}
        if var_name not in truth_map or var_name not in var_values:
            raise ValueError(
                f"Variable {var_name} not found in truth_map or var_values"
            )
        return var_values[var_name].strip() in [
            v.strip() for v in truth_map[var_name].split("|")
        ]

    # Create a map of variable values (True/False)
    var_map = {}
    unique_vars = set(struct_vars)
    for var_name in unique_vars:
        if var_name not in truth_map or var_name not in var_values:
            raise ValueError(
                f"Variable {var_name} not found in truth_map or var_values: {truth_map} {var_values}"
            )
        allowed = [v.strip() for v in truth_map[var_name].split("|")]
        actual = var_values[var_name]
        var_map[var_name] = actual in allowed

    # Custom expression parser (no eval)
    def parse_expression(expr):
        # Remove all whitespace
        expr = re.sub(r"\s+", "", expr)

        # Handle parentheses first
        while "(" in expr:
            start = expr.rfind("(")
            end = expr.find(")", start)
            if end == -1:
                break
            inner_expr = expr[start + 1 : end]
            result = parse_expression(inner_expr)
            expr = expr[:start] + ("true" if result else "false") + expr[end + 1 :]

        # Handle NOT operations
        expr = expr.replace("!true", "false").replace("!false", "true")

        # Handle AND operations
        while "&" in expr:
            parts = expr.split("&")
            result = all(part == "true" for part in parts)
            expr = str(result).lower()

        # Handle OR operations
        while "|" in expr:
            parts = expr.split("|")
            result = any(part == "true" for part in parts)
            expr = str(result).lower()

        return expr == "true"

    # Replace {variable} tokens with their boolean values
    processed_expr = struct_expr
    for var_name in unique_vars:
        regex = re.compile(r"\{" + re.escape(var_name) + r"\}")
        processed_expr = regex.sub(str(var_map[var_name]).lower(), processed_expr)

    # Parse the expression
    try:
        return parse_expression(processed_expr)
    except Exception as e:
        print("Error evaluating condition:", e, "Expression:", processed_expr)
        return False


# Example usage
if __name__ == "__main__":
    struct_expr = "(!{var1}|!{var2})&({var1}|{var3})"

    truth_map = {
        "var1": "Collective action|Individual action",
        "var2": "Yes|Partly",
        "var3": "No",
    }

    var_values = {
        "var1": "Individual action",
        "var2": "No",
        "var3": "No",
    }

    assert condition_is_active(struct_expr, truth_map, var_values) == True  # ✅ True
    assert (
        condition_is_active("({var1}&{var2})&({var1}|{var3})", truth_map, var_values)
        == False
    )
    assert condition_is_active("{var1}", truth_map, var_values) == True
    assert condition_is_active("!{var1}", truth_map, var_values) == False
    # Test bare name (no braces) — backward compat for simple single-variable case
    assert condition_is_active("var1", truth_map, var_values) == True
    print("All tests passed")
