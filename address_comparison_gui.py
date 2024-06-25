import os
import pandas as pd
from tkinter import Tk, Label, Button, Entry, filedialog, StringVar, messagebox, OptionMenu
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

def compare_addresses_fuzzy(list1, list2, threshold=80):
    matches = []
    for address1 in list1:
        best_match = process.extractOne(address1, list2, scorer=fuzz.token_set_ratio)
        if best_match[1] >= threshold:
            matches.append((address1, best_match[0], best_match[1]))
    return matches

def cosine_similarity_score(str1, str2):
    vectorizer = CountVectorizer().fit([str1, str2])
    vectors = vectorizer.transform([str1, str2])
    cosine_sim = cosine_similarity(vectors)
    return cosine_sim[0][1]

def compare_addresses(list_a, list_b):
    similarity_scores = []
    similar_addresses = []
    for address_a in list_a:
        max_score = 0
        similar_address = None
        for address_b in list_b:
            score = cosine_similarity_score(address_a, address_b)
            if score > max_score:
                max_score = score
                similar_address = address_b
        similarity_scores.append(max_score*100)
        similar_addresses.append(similar_address)
    return similarity_scores, similar_addresses

def co_sine_similarity(old_address_df, new_address_df):
    scores, similar_addresses = compare_addresses(old_address_df, new_address_df)
    return old_address_df, scores, similar_addresses

def select_old_file():
    global old_file_path, df_old_addresses
    old_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
    if old_file_path:
        try:
            dataframe = pd.read_excel(old_file_path)
            column_name = column_entry.get().strip()
            if column_name not in dataframe.columns:
                messagebox.showerror("Error", f"Column '{column_name}' not found in the file.")
                return
            df_old_addresses = dataframe[column_name].tolist()
            messagebox.showinfo("Success", f"Old addresses file '{os.path.basename(old_file_path)}' selected successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Error opening file: {e}")

def select_new_file():
    global new_file_path, df_new_addresses
    new_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
    if new_file_path:
        try:
            dataframe = pd.read_excel(new_file_path)
            column_name = column_entry.get().strip()
            if column_name not in dataframe.columns:
                messagebox.showerror("Error", f"Column '{column_name}' not found in the file.")
                return
            df_new_addresses = dataframe[column_name].tolist()
            messagebox.showinfo("Success", f"New addresses file '{os.path.basename(new_file_path)}' selected successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Error opening file: {e}")

def run_comparison():
    global old_file_path, df_old_addresses, new_file_path, df_new_addresses
    
    if not old_file_path or not df_old_addresses:
        messagebox.showerror("Error", "Old addresses file not selected.")
        return

    if not new_file_path or not df_new_addresses:
        messagebox.showerror("Error", "New addresses file not selected.")
        return

    old_file_name = os.path.splitext(os.path.basename(old_file_path))[0]
    new_file_name = os.path.splitext(os.path.basename(new_file_path))[0]

    if algorithm.get() == "Co-sine Similarity":
        old_addresses, scores, new_addresses = co_sine_similarity(df_old_addresses, df_new_addresses)
        unmatched_new_addresses = set(df_new_addresses) - set(new_addresses)

        for address in unmatched_new_addresses:
            old_addresses.append("No Match")
            scores.append("NA")
            new_addresses.append(address)

        pairs = [{"Old Address": old_address, "Similarity Score": score, "New Address": new_address}
                 for old_address, score, new_address in zip(old_addresses, scores, new_addresses)]
        result_df = pd.DataFrame(pairs)
        output_file_name = f"Comparison_Co-sine_Similarity_{old_file_name}_vs_{new_file_name}.xlsx"

    elif algorithm.get() == "Fuzzy Wuzzy":
        old_addresses = []
        matched_addresses = []
        matched_score_FuzzyWuzzy = []
        results = compare_addresses_fuzzy(df_old_addresses, df_new_addresses)
        for match in results:
            old_addresses.append(match[0])
            matched_addresses.append(match[1])
            matched_score_FuzzyWuzzy.append(match[2])

        unmatched_new_addresses = set(df_new_addresses) - set(matched_addresses)

        for address in unmatched_new_addresses:
            old_addresses.append("No Match")
            matched_score_FuzzyWuzzy.append("NA")
            matched_addresses.append(address)

        pairs = [{"Old Address": old_address, "Similarity Score": score, "New Address": new_address}
                 for old_address, score, new_address in zip(old_addresses, matched_score_FuzzyWuzzy, matched_addresses)]
        result_df = pd.DataFrame(pairs)
        output_file_name = f"Comparison_FuzzyWuzzy_{old_file_name}_vs_{new_file_name}.xlsx"
    else:
        messagebox.showerror("Error", "Invalid algorithm selected.")
        return

    result_df.to_excel(output_file_name, index=False)
    messagebox.showinfo("Success", f"File saved successfully as {output_file_name}!")

root = Tk()
root.title("Address Comparison Tool")

old_file_path = None
new_file_path = None
df_old_addresses = None
df_new_addresses = None

Label(root, text="Old Addresses File:").grid(row=0, column=0, padx=10, pady=5)
Button(root, text="Select File", command=select_old_file).grid(row=0, column=1, padx=10, pady=5)

Label(root, text="New Addresses File:").grid(row=1, column=0, padx=10, pady=5)
Button(root, text="Select File", command=select_new_file).grid(row=1, column=1, padx=10, pady=5)

Label(root, text="Address Column Name:").grid(row=2, column=0, padx=10, pady=5)
column_entry = Entry(root)
column_entry.grid(row=2, column=1, padx=10, pady=5)

Label(root, text="Algorithm:").grid(row=3, column=0, padx=10, pady=5)
algorithm = StringVar(root)
algorithm.set("Co-sine Similarity")
OptionMenu(root, algorithm, "Co-sine Similarity", "Fuzzy Wuzzy").grid(row=3, column=1, padx=10, pady=5)

Button(root, text="Run Comparison", command=run_comparison).grid(row=4, column=0, columnspan=2, pady=20)

root.mainloop()
