def find_longest_common_prefix(given_words: list) -> str:
    """
    Given a list words, find the longest common prefix within array.

    :param given_words: names
    :return: common prefix
    """
    if not given_words:
        return ""
    # since list of string will be sorted and retrieved min max by alphabetic order
    s1 = min(given_words)
    print(f"s1: {s1}")
    s2 = max(given_words)
    print(f"s2: {s2}")

    for i, c in enumerate(s1):
        if c != s2[i]:
            return s1[:i]  # stop until hit the split index
    return s1


names: list = ["Bill", "Bob", "Billy"]

for name in names:
    print(name)


# foo = find_longest_common_prefix(names)
# print(foo)
