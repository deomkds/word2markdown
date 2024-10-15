import subprocess
import os
import time


def list_files_in(path):
    word_docs = []

    for root, dirs, files in os.walk(path, topdown=False):
        for name in files:
            if name.endswith(".doc") or name.endswith(".docx") or name.endswith(".rtf") or name.endswith(".md"):
                word_docs.append(os.path.join(root, name))

    return word_docs


def count_markdowns(path):
    count = 0
    for root, dirs, files in os.walk(path, topdown=False):
        for name in files:
            if name.endswith(".md"):
                count += 1

    return count


def file_name_without_extension(file_path):
    file_name = os.path.basename(file_path)
    return file_name[:file_name.find(".")].strip()


def timestamp_of_file(file_path):
    file_modification_time = os.path.getmtime(file_path)       # In seconds since epoch.
    timestamp_string = time.ctime(file_modification_time)      # As a timestamp string.
    time_object = time.strptime(timestamp_string)              # To a timestamp object.
    return time.strftime("%Y %m %d", time_object)              # To my format.


def set_original_timestamp(old_file_path, new_file_path):
    modified_time = os.path.getmtime(old_file_path)
    accessed_time = os.path.getatime(old_file_path)
    os.utime(new_file_path, (accessed_time, modified_time))


def convert_to_markdown(input_path, output_path):
    saida = subprocess.Popen([
        "pandoc",
        "-t",
        "markdown-escaped_line_breaks+backtick_code_blocks+grid_tables ",
        "--wrap=none",
        "--reference-links",
        "-o",
        output_path,
        input_path
    ], stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    std_out = str(saida.communicate()[0])
    std_err = str(saida.communicate()[1])

    print(f"Pandoc stdout: {std_out}")
    print(f"Pandoc stderr: {std_err}")

    output_file = os.path.basename(output_path)
    file_exists = os.path.exists(output_path)

    if file_exists:
        print(f"File {output_file} exists before cleaning.")
    else:
        raise FileNotFoundError(f"File {output_file} does not exist!!")

    with open(output_path, "r") as arquivo:
        data = arquivo.read()

    with open(output_path, "w") as arquivo:
        arquivo.write(data.replace("\\", ""))

    output_file = os.path.basename(output_path)
    file_exists = os.path.exists(output_path)

    if file_exists:
        print(f"File {output_file} exists after cleaning.")
    else:
        raise FileNotFoundError(f"File {output_file} does not exist!!")

    set_original_timestamp(input_path, output_path)


src_path = "/home/deomkds/Desktop/Teste/"
list_of_files = list_files_in(src_path)

total_files = len(list_of_files)
print(f"About to convert {total_files} files.\n")

for index, file_path in enumerate(list_of_files):
    input_file = file_path
    print(f"Processing file {index + 1} of {total_files}: {os.path.basename(input_file)}")
    convert_to_markdown(input_file, file_path.replace(".docx", ".md"))

print(f"Deleting Word files: ", end="")

for file_path in list_of_files:
    print(".", end="")
    os.remove(file_path)

print(f"\n\nConverted {count_markdowns(src_path)} files.")
