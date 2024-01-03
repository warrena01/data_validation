from transformers import BertModel, BertTokenizer
import torch
from scipy.spatial.distance import cosine
import difflib


def get_bert_embedding(text, model, tokenizer):
    tokens = tokenizer.tokenize(tokenizer.decode(tokenizer.encode(text, add_special_tokens=True)))
    inputs = tokenizer.encode(text, return_tensors="pt")
    outputs = model(inputs)
    return outputs.last_hidden_state.mean(dim=1).squeeze().detach().numpy()

def bert_similarity_score(str1, str2, model, tokenizer):
    embedding1 = get_bert_embedding(str1, model, tokenizer)
    embedding2 = get_bert_embedding(str2, model, tokenizer)
    return 1 - cosine(embedding1, embedding2)

def camel_case(s):
    """
    Convert a string to title case.
    """
    parts = s.split()
    return ' '.join(part.capitalize() for part in parts)

def main():
    # Load pre-trained BERT model and tokenizer
    model_name = "bert-base-uncased"
    model = BertModel.from_pretrained(model_name)
    tokenizer = BertTokenizer.from_pretrained(model_name)

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
        most_similar, score = max(((item, bert_similarity_score(i, item, model, tokenizer)) for item in rms_list), key=lambda x: x[1], default=("", 0))
        output_dict[i] = {"most_similar": most_similar, "similarity_score": score}

    for i in new_input_list:
        print(i, ':', output_dict[i]["most_similar"], output_dict[i]["similarity_score"])

if __name__ == "__main__":
    main()
