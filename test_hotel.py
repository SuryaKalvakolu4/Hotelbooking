import unittest
from tkinter import Tk
from Hotel import HotelBookingApp

class TestHotelBookingApp(unittest.TestCase):
    def setUp(self):
        self.root = Tk()
        self.app = HotelBookingApp(self.root)

    def tearDown(self):
        self.root.destroy()

    def test_book_room(self):
        # Test booking a room
        self.app.user_entry.insert(0, "John")
        self.app.room_entry.insert(0, "101")
        self.app.book_room()
        bookings = self.app.sheet.iter_rows(values_only=True)
        self.assertTrue(any(row[0] == 101 and row[1] == "John" for row in bookings))

    def test_view_bookings(self):
        # Test viewing bookings
        self.app.view_bookings()
        self.assertTrue(self.app.sheet.max_row > 0)

    def test_delete_booking(self):
        # Test deleting a booking
        self.app.room_entry.insert(0, "101")
        self.app.delete_booking()
        bookings = self.app.sheet.iter_rows(values_only=True)
        self.assertFalse(any(row[0] == 101 for row in bookings))

if __name__ == '__main__':
    unittest.main()
