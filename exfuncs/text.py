from re import sub
import string


def camel_case(i_string):
    """
    Convert input string to camelCase.
    """
    if i_string is None:
        return ""
    if len(i_string) <= 1:
        return i_string

    c_string = sub(r"(_|-)+", " ", i_string).title().replace(" ", "").replace("\n", "")
    return c_string[0].lower() + c_string[1:]


def remove_punctuation(i_string):
    """
    Return new string with punctuation removed from input string.
    """
    table = str.maketrans(dict.fromkeys(string.punctuation))
    new_s = i_string.translate(table)
    return new_s


def remove_parens(i_string):
    """
    Returns new string with parentheses and content inside parentheses removed.
    """
    return sub(r" ?\([^)]+\)", "", i_string)
