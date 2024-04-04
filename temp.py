
# Variant Soni
variant_soni_label2 = tk.Label(frame2, text="Variant Soni:")
variant_soni_label2.grid(row=5, column=0, padx=10, pady=5)
variant_soni_entry2 = tk.Entry(frame2)
variant_soni_entry2.grid(row=5, column=1, padx=10, pady=5)

# Variant Test Soni
variant_test_soni_label2 = tk.Label(frame2, text="Testlar Soni:")
variant_test_soni_label2.grid(row=6, column=0, padx=10, pady=5)
variant_test_soni_entry2 = tk.Entry(frame2)
variant_test_soni_entry2.grid(row=6, column=1, padx=10, pady=5)

# Til
til_label2 = tk.Label(frame2, text="Til:")
til_label2.grid(row=0, column=0, padx=5, pady=5)
til_var = tk.StringVar()
til_options = ["uz", "ru"]
for index, til in enumerate(til_options):
    tk.Radiobutton(frame2, text=til, variable=til_var, value=til).grid(row=0, column=index + 1, padx=10)

# Kafedra nomi
kafedra_nomi_label2 = tk.Label(frame2, text="Kafedra nomi:")
kafedra_nomi_label2.grid(row=1, column=0, padx=10, pady=5,  sticky="W")
kafedra_nomi_entry2 = tk.Entry(frame2)
kafedra_nomi_entry2.grid(row=1, column=1, padx=10, pady=5)

# Guruh nomi
guruh_nomi_label2 = tk.Label(frame2, text="Guruh:")
guruh_nomi_label2.grid(row=3, column=0, padx=10, pady=5, sticky="W")
guruh_nomi_entry2 = tk.Entry(frame2)
guruh_nomi_entry2.grid(row=3, column=1, padx=10, pady=5)

# Fan nomi
fan_nomi_label2 = tk.Label(frame2, text="Fan nomi:")  # Added Fan nomi
fan_nomi_label2.grid(row=2, column=0, padx=10, pady=5,  sticky="W")  # Added Fan nomi
fan_nomi_entry2 = tk.Entry(frame2)  # Added Fan nomi
fan_nomi_entry2.grid(row=2, column=1, padx=10, pady=5)  # Added Fan nomi

# Semestr
semestr_label2 = tk.Label(frame2, text="Semestr:")
semestr_label2.grid(row=4, column=0, padx=10, pady=5,sticky="W")
semestr_entry2 = tk.Entry(frame2)
semestr_entry2.grid(row=4, column=1, padx=10, pady=5)

# Tuzuvchi
tuzuvchi_label2 = tk.Label(frame2, text="Tuzuvchi:")
tuzuvchi_label2.grid(row=7, column=0, padx=10, pady=5, sticky="W")
tuzuvchi_entry2 = tk.Entry(frame2)
tuzuvchi_entry2.grid(row=7, column=1, padx=10, pady=5)

# Zavod Kafedra
zav_kaf_label2 = tk.Label(frame2, text="Kafedra mudiri:")
zav_kaf_label2.grid(row=8, column=0, padx=10, pady=5, sticky="W")
zav_kaf_entry2 = tk.Entry(frame2)
zav_kaf_entry2.grid(row=8, column=1, padx=10, pady=5)

# Button to open file dialog for file path
open_file_button = tk.Button(frame2, text="Open File", command=open_file_path)
open_file_button.grid(row=9, column=0, pady=10)

# Entry for file path
file_path_entry2 = tk.Entry(frame2)
file_path_entry2.grid(row=9, column=1, padx=10, pady=5)

# "Tayyorla" button
tayyorla_button2 = tk.Button(frame2, text="Tayyorla", command=tayyorla_clicked, height=1)
tayyorla_button2.grid(row=10, column=0, padx=10, pady=5)