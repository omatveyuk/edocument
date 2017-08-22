""" Cosine similarity algorithm.
    Define semilarity of two strings.
    cosine_semilarity = dot_product/ magnitude1*magnitude2

# >>> cosine_semilarity("I love Uliana", "I love Uliana")
# 1.0

# >>> cosine_semilarity("I love Uliana", "We")
# 0.0

>>> cosine_semilarity("I love Uliana", "I love Anastasia")
0.67

# >>> s1 = "Its a penalty goal .In football a single score can change game. what a game. Players are dancing"
# >>> s2 = "In football with penalty you can score a goal that can change the game. what a game. Players are dancing "
# >>> cosine_semilarity(s1, s2)
# 0.79

"""

import math

def cosine_semilarity(str1, str2):
    """Given two strings, return value of cosine semilarity."""

    word1 = str1.lower().split(' ')
    word2 = str2.lower().split(' ')
    dict_words = {}

    for word in word1:
        if word in dict_words:
            dict_words[word]["str1"] = dict_words[word]["str1"] + 1
        else:
            dict_words[word] = {"str1": 1, "str2": 0}

    for word in word2:
        if word in dict_words:
            dict_words[word]["str2"] = dict_words[word]["str2"] + 1
        else:
            dict_words[word] = {"str1": 0, "str2": 1}

    dot_product = 0
    magnitude1 = 0
    magnitude2 = 0
    for key in dict_words:
        dot_product += dict_words[key]["str1"] * dict_words[key]["str2"]
        magnitude1 += dict_words[key]["str1"] ** 2
        magnitude2 += dict_words[key]["str2"] ** 2

    cos_semilarity = round(dot_product / (math.sqrt(magnitude1) * math.sqrt(magnitude2)), 2)
    return cos_semilarity


if __name__ == '__main__':
    import doctest
    if doctest.testmod().failed == 0:
        print "\n*** ALL TEST PASSED. W00T! ***\n"