"""
LC practice.
"""


class Solution:
    """
    Solution to easy problem.
    """
    def __init__(self, numbers: list, target: int):
        self.numbers: list = numbers
        self.target: int = target

    def main(self):
        equal_targets: list = []

        for index_1, value_1 in enumerate(self.numbers, start=0):
            for index_2, value_2 in enumerate(self.numbers, start=0):

                if index_1 == index_2:
                    pass
                elif self.is_numbers_equal(value_1, value_2):
                    equal_targets.append(value_1)
                    equal_targets.append(value_2)

        return list(set(equal_targets))

    def is_numbers_equal(self, number_1: int, number_2: int) -> bool:
        """
        Checks if numbers are equal to the target when combined.
        """
        total: int = number_1 + number_2

        if self.target == total:
            return True

        elif self.target != total:
            return False


example_numbers: list = [1, 2, 3, 4, 5]
example_target: int = 5

foo = Solution(example_numbers, example_target)
bar = foo.main()
print(bar)