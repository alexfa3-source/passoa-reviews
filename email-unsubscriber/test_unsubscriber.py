import unittest
import unsubscriber as u

PROMO = (b"From: =?UTF-8?B?zqPOv8+Gzq/OsQ==?= <news@shop.gr>\r\n"
         b"Subject: =?UTF-8?Q?Prosfores?=\r\n"
         b"Date: Sat, 31 Oct 2026 10:00:00 +0200\r\n"
         b"List-Unsubscribe: <mailto:unsub@shop.gr?subject=stop>,\r\n <https://shop.gr/u?id=1>\r\n"
         b"List-Unsubscribe-Post: List-Unsubscribe=One-Click\r\n\r\n")
OLDER = PROMO.replace(b"31 Oct", b"20 Oct").replace(b"id=1", b"id=old")
PERSONAL = b"From: Maria <maria@example.com>\r\nSubject: hi\r\nDate: Sat, 31 Oct 2026 10:00:00 +0200\r\n\r\n"
MAILTO_ONLY = b"From: x@list.org\r\nList-Unsubscribe: <mailto:leave@list.org>\r\n\r\n"


class T(unittest.TestCase):
    def test_analyze(self):
        s = u.analyze_headers([OLDER, PROMO, PERSONAL, MAILTO_ONLY])
        self.assertEqual(set(s), {"news@shop.gr", "x@list.org"})
        shop = s["news@shop.gr"]
        self.assertEqual(shop.count, 2)
        self.assertEqual(shop.name, "Σοφία")
        self.assertEqual(shop.http_links, ["https://shop.gr/u?id=1"])  # πιο πρόσφατο
        self.assertEqual(shop.method, "one-click")
        self.assertEqual(s["x@list.org"].method, "e-mail")

    def test_selection(self):
        self.assertEqual(u.parse_selection("1,3, 5-7", 6), {1, 3, 5, 6})
        self.assertEqual(u.parse_selection("all", 3), {1, 2, 3})


if __name__ == "__main__":
    unittest.main()
