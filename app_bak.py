import pandas as pd
import random
import tkinter as tk
import sys
import os

excel_file_path = os.path.join(os.path.dirname(sys.executable), "単語テストツール", "word_bank.xlsx")
df = pd.read_excel(excel_file_path, sheet_name='日常用语')

window = tk.Tk()

word_history = []
current_index = -1
correct_word = ""
correct_word_kana = ""

def display_random_word():
    global current_index, correct_word, correct_word_kana
    entry.delete(0, 'end')
    result_label.config(text="", bg="white")

    start = range_start_entry.get()
    end = range_end_entry.get()

    if start == '' or end == '':
        available_rows = df
    else:
        start = int(start)
        end = int(end)

        available_rows = df[df['No.'].between(start, end)]
        if available_rows.empty:
            result_label.config(text="No words in the specified range.", bg="light coral")
            return
    available_words = available_rows['中文'].tolist()
    unused_words = [word for word in available_words if word not in word_history]

    if not unused_words:
        result_label.config(text="All words in the range have been used. Resetting history.", bg="light coral")
        word_history.clear()

    while True:
        if unused_words:
            word = random.choice(unused_words)
            break
        else:
            word = random.choice(available_words)

        if word not in word_history:
            break
    # random_row = available_rows.sample()
    # word = random_row['中文'].values[0]
    label.config(text=word)
    random_row = available_rows[available_rows['中文'] == word]
    correct_word = random_row['日文'].values[0]
    correct_word_kana = random_row['日文（ひらがな）'].values[0]
    word_history.append(word)
    current_index = len(word_history) - 1

def show_previous_word():
    global current_index, correct_word, correct_word_kana
    if current_index > 0:
        current_index -= 1
        label.config(text=word_history[current_index])
        correct_row = df[df['中文'] == word_history[current_index]]
        correct_word = correct_row['日文'].values[0]
        correct_word_kana = correct_row['日文（ひらがな）'].values[0]

def check_input(event):
    user_input = entry.get()
    
    if user_input == correct_word or user_input == correct_word_kana:
        result_label.config(text="Correct!", bg="light green")
    else:
        result_label.config(text="Incorrect!", bg="light coral")

def update_word_range():
    display_random_word()

def show_tips():
    if correct_word and correct_word_kana:
        result_label.config(text=f"Correct Word: {correct_word}\nCorrect Word (Kana): {correct_word_kana}")
    else:
        result_label.config(text="No word to display tips for.")

label = tk.Label(window, text="", font=("Helvetica", 16))
label.pack()

entry = tk.Entry(window, width=100)
entry.bind("<KeyRelease>", check_input)
entry.pack()

range_frame = tk.Frame(window)
range_frame.pack()

range_label = tk.Label(range_frame, text="Range: from")
range_label.pack(side="left")

range_start_entry = tk.Entry(range_frame, width=5)
range_start_entry.pack(side="left")

range_label_to = tk.Label(range_frame, text="to")
range_label_to.pack(side="left")

range_end_entry = tk.Entry(range_frame, width=5)
range_end_entry.pack(side="left")

prev_button = tk.Button(window, text="Previous", command=show_previous_word)
prev_button.pack()

next_button = tk.Button(window, text="Next", command=display_random_word)
next_button.pack()

update_button = tk.Button(window, text="Update Range", command=update_word_range)
update_button.pack()

result_label = tk.Label(window, text="", bg="white")
result_label.pack()

tips_button = tk.Button(window, text="Tips", command=show_tips)
tips_button.pack()

display_random_word()

window.mainloop()