# Definition for a binary tree node.
class TreeNode:
    def __init__(self, x):
        self.val = x
        self.left = None
        self.right = None


# inorder = [9,3,15,20,7]
# postorder = [9,15,7,20,3]


class Solution:
    def buildTree(self, inorder, postorder):
        """
        :type inorder: List[int]
        :type postorder: List[int]
        :rtype: TreeNode
        """
        if len(postorder) == 0:
            return None
        res = TreeNode(postorder[-1])
        index = inorder.index(res.val)
        res.left = self.buildTree(inorder[0:index], postorder[0:index])
        res.right = self.buildTree(inorder[index + 1:], postorder[index:-1])

        return res
