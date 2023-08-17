def generate_emails(input_string):
    emails = set()
    email = input_string.replace(' ', '') + '@gmail.com'
    emails.add(email)

    # add dots after every character
    for i in range(len(input_string)):
        if input_string[i] != ' ':
            email = input_string[:i] + '.' + input_string[i:].replace(' ', '') + '@gmail.com'
            emails.add(email)
            print(email)

    # add dots after every two characters
    for i in range(len(input_string)):
        if input_string[i] != ' ':
            for j in range(i + 1, len(input_string)):
                if input_string[j] != ' ':
                    email = input_string[:i] + '.' + input_string[i:j] + '.' + input_string[j:].replace(' ',
                                                                                                        '') + '@gmail.com'
                    emails.add(email)
                    print(email)

    # add dots after every three characters
    for i in range(len(input_string)):
        if input_string[i] != ' ':
            for j in range(i + 1, len(input_string)):
                if input_string[j] != ' ':
                    for k in range(j + 1, len(input_string)):
                        if input_string[k] != ' ':
                            email = input_string[:i] + '.' + input_string[i:j] + '.' + input_string[
                                                                                       j:k] + '.' + input_string[
                                                                                                    k:].replace(' ',
                                                                                                                '') + '@gmail.com'
                            emails.add(email)
                            print(email)

    return emails




string = "mymailautt"
variations = []

# generate all possible combinations of dot positions using binary numbers
for i in range(2**(len(string)-2)):
    binary = bin(i)[2:].zfill(len(string)-2)
    new_string = string[0]
    for j in range(len(binary)):
        if binary[j] == "1":
            new_string += "."
        new_string += string[j+1]
    variations.append(new_string)
    print(new_string+"@gmail.com,Vodafone_044834")

print(variations)

# input_string = "myamilaut"
# emails = generate_emails(input_string)
# print(emails)
