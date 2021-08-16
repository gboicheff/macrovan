
def make_freq_dict(entries):
    freq_dict = {}
    for entry in entries:
        if entry in freq_dict:
            freq_dict[entry] +=1
        else:
            freq_dict[entry] = 1
    return freq_dict


def find_duplicates(freq_dict):
    duplicate_list = []
    for (entry, freq) in freq_dict.items():
        if freq > 1:
            duplicate_list += (entry, freq)
    return duplicate_list

with open("entries.txt") as file:
    file_contents = file.readlines()
    freq_dict = make_freq_dict(file_contents)
    print(find_duplicates(freq_dict))


