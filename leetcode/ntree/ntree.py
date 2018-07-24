class Node(object):
    def __init__(self, val, children):
        self.val = val
        self.children = children


# Definition for a binary tree node.
class TreeNode(object):
    def __init__(self, x):
        self.val = x
        self.left = None
        self.right = None


class Codec:
    def encodeHelper(self, rootlist, i):
        if len(rootlist) == 0:
            return None
        if i == len(rootlist) - 1:
            res = TreeNode(rootlist[i].val)
            res.right = self.encodeHelper(rootlist[i].children, 0)
            return res
        else:
            res = TreeNode(rootlist[i].val)
            res.left = self.encodeHelper(rootlist, i + 1)
            res.right = self.encodeHelper(rootlist[i].children, 0)
            return res

    def encode(self, root):
        """Encodes an n-ary tree to a binary tree.
        :type root: Node
        :rtype: TreeNode
        """
        if not root:
            return None

        return self.encodeHelper([root], 0)

    def decodeHelper(self, res, data, currentlist):
        if not data:
            return res

        if data.right and not data.left:
            tmp = Node(data.right.val, [])
            tmp = self.decodeHelper(tmp, data.right, [tmp])
            res.children.append(tmp)
            return res

        if data.left and not data.right:
            tmp = Node(data.left.val, [])
            currentlist.append(tmp)
            tmp = self.decodeHelper(tmp, data.left, currentlist)
            return res

        if data.right and data.left:
            tmp = Node(data.left.val, [])
            currentlist.append(tmp)
            tmp = self.decodeHelper(tmp, data.left, currentlist)
            tmp = Node(data.right, [])
            res.children.append(tmp)
            tmp = self.decodeHelper(tmp, data.right, [tmp])
            return res

    def decode(self, data):
        """Decodes your binary tree to an n-ary tree.
        :type data: TreeNode
        :rtype: Node
        """
        if not data:
            return None

        res = Node(data.val, [])

        res = self.decodeHelper(res, data, [res])

        return res


class Solution(object):
    def levelOrder(self, root):
        """
        :type root: Node
        :rtype: List[List[int]]
        """
        if not root:
            return []

        res = []
        queue = []
        queue.append(root)
        while len(queue) > 0:
            subres = []
            tmpqueue = []
            for element in queue:
                if element:
                    subres.append(element.val)
                    if element.children:
                        for i in range(len(element.children)):
                            tmpqueue.append(element.children[i])

            queue = tmpqueue
            res.append(subres)
        return res


# Your Codec object will be instantiated and called as such:
codec = Codec()
root5 = Node(5, [])
root6 = Node(6, [])
root2 = Node(2, [])
root4 = Node(4, [])
root3 = Node(3, [root5, root6])
root1 = Node(1, [root3, root2, root4])

a = Solution()

res = codec.decode(codec.encode(root1))

ano = a.levelOrder(res)

print(ano)
