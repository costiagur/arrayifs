# arrayifs
VBA function for Excel that returns array of values based on conditions. Afterwards, user can do any operation available on an array.

Agruments:
valrange - values to be returned in an array.
testrange (1 to 4)- ranges of values to be tested with the conditions. First testrange is obligatory, others are optional.
condition (1 tp 4) - conditions. First condition is obligatory, others are optional.
ifsort - Receives True or False. Tells whether or not to perform soring of results in ascending order.

Returned result is an array. Afterwards, user can sum, count, max, min or apply other functions to that array.
