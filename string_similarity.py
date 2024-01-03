import difflib

def similarity_score(str1, str2):
    """
    Calculate a similarity score between two strings using SequenceMatcher.
    """
    matcher = difflib.SequenceMatcher(None, str1, str2)
    return matcher.ratio()

def camel_case(s):
    """
    Convert a string to title case.
    """
    parts = s.split()
    return ' '.join(part.capitalize() for part in parts)

def main():

    input_list = ["café", "résumé", "piñata", "jalapeño"]
    rms_list = ["cafe", "resume", "pinata", "jalapeno"]
    output_dict = {}
    input_list_unique = list(set(input_list))

    if len(input_list) == 0:
        print("Input lists should not be empty.")
        return

    input_list_unique = [camel_case(word) for word in input_list_unique]
    rms_list = [camel_case(word) for word in rms_list]

    new_input_list = []

    for i in input_list_unique:
        if i not in rms_list:
            new_input_list.append(i)

    for i in new_input_list:
        most_similar, score = max(((item, similarity_score(i, item)) for item in rms_list), key=lambda x: x[1], default=("", 0))
        output_dict[i] = {"most_similar": most_similar, "similarity_score": score}

    for i in new_input_list:
        print(i, ':', output_dict[i]["most_similar"], output_dict[i]["similarity_score"])

if __name__ == "__main__":
    main()
