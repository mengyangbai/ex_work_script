import unittest
import filter


class TestFilter(unittest.TestCase):
    '''test

    Arguments:
        unittest {[filter]} -- [emmis_get_box_no]
    '''

    def test_emmis_get_box_no(self):
        '''Get all finished box_no

        Arguments:
            box_no {list} -- all box no that has invoice
        Return:
            res {set} -- all box no that is finished in emmis
        '''
        self.assertEqual(filter.emmis_get_box_no(['MHT82531']), set())
        self.assertEqual(
            filter.emmis_get_box_no(['MHT56247']), set(['MHT56247']))
        self.assertEqual(filter.emmis_get_box_no([]), set())


if __name__ == '__main__':
    unittest.main()
