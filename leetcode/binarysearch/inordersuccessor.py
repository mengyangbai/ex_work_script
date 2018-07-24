# Definition for a binary tree node.
class TreeNode(object):
    def __init__(self, x):
        self.val = x
        self.left = None
        self.right = None


class Solution(object):
    def helper(self, root):
        res = []
        if not root:
            return None
        res.append(self.helper(root.left))
        res.append(root.val)
        res.append(self.helper(root.right))
        return res

    def inorderSuccessor(self, root, p):
        """
        :type root: TreeNode
        :type p: TreeNode
        :rtype: TreeNode
        """
        res = self.helper(root)
        return res[res.index(p) + 1]


b = TreeNode(0)

a = Solution()
print(a.inorderSuccessor(b, 0))
