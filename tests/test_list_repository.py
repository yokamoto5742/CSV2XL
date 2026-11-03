import sqlite3

from services.list_repository import EXCLUDE_DOCS, STAFF, ListRepository


class TestListRepository:
    def test_empty(self, connection: sqlite3.Connection) -> None:
        assert ListRepository(connection).get_items(STAFF) == []

    def test_round_trip_keeps_order(self, connection: sqlite3.Connection) -> None:
        repository = ListRepository(connection)

        repository.set_items(STAFF, ['山本', '植田', '沖野'])

        assert repository.get_items(STAFF) == ['山本', '植田', '沖野']

    def test_strips_blank_and_duplicates(self, connection: sqlite3.Connection) -> None:
        repository = ListRepository(connection)

        repository.set_items(STAFF, [' 山本 ', '', '山本', '植田'])

        assert repository.get_items(STAFF) == ['山本', '植田']

    def test_replaces_only_same_category(self, connection: sqlite3.Connection) -> None:
        repository = ListRepository(connection)
        repository.set_items(STAFF, ['山本'])
        repository.set_items(EXCLUDE_DOCS, ['紹介状'])

        repository.set_items(STAFF, ['植田'])

        assert repository.get_items(STAFF) == ['植田']
        assert repository.get_items(EXCLUDE_DOCS) == ['紹介状']
