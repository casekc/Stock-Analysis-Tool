import pandas as pd




def format_list(financials_list):
    numbers_list = []
    for i in financials_list:
        numbers = i[10:].split('$')[1:]
        numbers_list.append(numbers)
    return numbers_list



def move_hyphen(nested_list):
    result = []
    for sublist in nested_list:
        new_sublist = []
        for element in sublist:
            if element.endswith("-"):
                element = element[:-1]  # Remove the hyphen
                if new_sublist and new_sublist[-1].endswith("-"):
                    new_sublist[-1] += element  # Append the element to the last element in the new sublist
                else:
                    new_sublist.append(
                        element)  # Append the element as is if the new sublist is empty or the last element doesn't end with a hyphen
                new_sublist.append("-")  # Append the hyphen as a separate element
            else:
                new_sublist.append(element)
        result.append(new_sublist)
    return result

def split_list(data):
    sublists = []
    sublist = []
    total_revenue_count = 0

    for item in data:
        if item == "Total Revenue":
            total_revenue_count += 1
            if total_revenue_count > 2:
                sublists.append(sublist)
                sublist = []
        sublist.append(item)

    # Append the last sublist
    sublists.append(sublist)

    return sublists

def extract_net_income(data):
    extracted_data = []
    for sublist in data:
        if 'Net Income' in sublist:
            index = sublist.index('Net Income')
            extracted_data.append(sublist[index:index+5])
    return extracted_data


''' 
output1 = format_list(string)
output2 = move_hyphen(output1)
for sublist in output2:
    index = None  # Initialize the index variable
    if '-' in sublist:
        index = sublist.index('-')
    if index is not None and index + 1 < len(sublist):
        sublist[index] = '-' + sublist[index + 1]
        del sublist[index + 1]
print(output2)
'''
# numbers_list = format_list(financials_list)
# output = move_hyphen(numbers_list)


# nested_list = output

# final_list = nested_list
