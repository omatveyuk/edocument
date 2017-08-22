""" Jaccard Similarity algorithm.
    Define semilarity of two strings.
    sij=p/(p+q+r)
    p = # of attributes in both objects 
    q = # of attributes only in 1st object 
    r = # of attributes only in 2nd onject
    You can count the repeated words as 1 word. (version 1)    
    You can count the repeated words as if they were different words. (vesrion 2) 
    Return tuple (version1, verion2)
    

>>> jaccard_semilarity("I love Uliana", "I love Uliana")
(1.0, 1.0)

>>> jaccard_semilarity("I love Uliana", "Uliana I love Uliana")
(1.0, 0.75)

>>> jaccard_semilarity("I love Uliana", "Anastasia I love Anastasia")
(0.5, 0.4)

>>> jaccard_semilarity("I love Uliana", "We")
(0.0, 0.0)


>>> jaccard_semilarity("I love Uliana", "I love Anastasia")
(0.5, 0.5)

# >>> s1 = "Its a penalty goal .In football a single score can change game. what a game. Players are dancing"
# >>> s2 = "In football with penalty you can score a goal that can change the game. what a game. Players are dancing "
# >>> jaccard_semilarity(s1, s2)

"""

import math

def jaccard_semilarity(str1, str2):
    """Given two strings, return value of jaccard semilarity."""

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

    p1 = 0
    q1 = 0
    r1 = 0

    p2 = 0
    q2 = 0
    r2 = 0

    for key in dict_words:
        p1 += 1 if (dict_words[key]["str1"] > 0 and dict_words[key]["str2"] > 0) else 0
        q1 += 1 if (dict_words[key]["str1"] > 0 and dict_words[key]["str2"] == 0) else 0
        r1 += 1 if (dict_words[key]["str1"] == 0 and dict_words[key]["str2"] > 0) else 0

        p2 += dict_words[key]["str2"] if dict_words[key]["str1"] > dict_words[key]["str2"] else dict_words[key]["str1"]
        q2 += (dict_words[key]["str1"] - dict_words[key]["str2"]) if dict_words[key]["str1"] > dict_words[key]["str2"] else 0
        r2 += (dict_words[key]["str2"] - dict_words[key]["str1"]) if dict_words[key]["str1"] < dict_words[key]["str2"] else 0

    jac_semilarity = (round(p1 * 1.0/ (p1+q1+r1), 2), round(p2 * 1.0/ (p2+q2+r2),2))
    return jac_semilarity


if __name__ == '__main__':
    import doctest
    if doctest.testmod().failed == 0:
        print "\n*** ALL TEST PASSED. W00T! ***\n"