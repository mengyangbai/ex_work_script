class Solution(object):
    def pivotIndex(self, nums):
        """
        :type nums: List[int]
        :rtype: int
        """
        sums = sum(nums)
        tmp_sum = 0
        for i in range(len(nums)):
            tmp_compare = tmp_sum * 2 + nums[i]
            if tmp_compare == sums:
                return i
            tmp_sum = tmp_sum + nums[i]

        return -1
