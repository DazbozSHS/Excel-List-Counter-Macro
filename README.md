# Excel-List-Counter-Macro

# A VBA macro to count the number of matching instances in a list column, then reduce the list to only unique values with original counts next to them. This macro also formats the list into a standardized style which can be edited as desired. 

# So for example a column list of
# apple
# apple
# orange
# will return two columns showing each unique value with it's count:
# apple   2
# orange  1

# RESTRICTIONS:
# The macro uses relative references, but the list must be in ROW A (any collumn), with the top header cell selected.

# HOW TO USE:
# Import to Excel VBA, activate with CTRL+SHIT+C
