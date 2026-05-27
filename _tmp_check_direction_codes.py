from __future__ import annotations

import sqlite3
from pathlib import Path


DB_PATH = (
    Path("02_Result")
    / "차종별_교통량_추출"
    / "엠큐닉_2601_2604"
    / "기후에너지제공데이터(260101~260430).db"
)


def main() -> None:
    connection = sqlite3.connect(DB_PATH.resolve().as_uri() + "?mode=ro", uri=True)
    connection.execute("PRAGMA query_only = ON")
    cursor = connection.cursor()

    print("target_codes")
    target_rows = list(
        cursor.execute(
            """
            SELECT direction, COUNT(*)
            FROM traffic_by_vehicle
            WHERE direction IN ('00', '07', '09', '11')
            GROUP BY direction
            ORDER BY direction
            """
        )
    )
    if target_rows:
        for direction, count in target_rows:
            print(f"{direction}\t{count}")
    else:
        print("none")

    print("numeric_only")
    numeric_rows = list(
        cursor.execute(
            """
            SELECT direction, COUNT(*)
            FROM traffic_by_vehicle
            WHERE direction IS NOT NULL
              AND direction != ''
              AND direction NOT GLOB '*[^0-9]*'
            GROUP BY direction
            ORDER BY direction
            """
        )
    )
    if numeric_rows:
        for direction, count in numeric_rows:
            print(f"{direction}\t{count}")
    else:
        print("none")

    null_count, blank_count = cursor.execute(
        """
        SELECT
            SUM(CASE WHEN direction IS NULL THEN 1 ELSE 0 END),
            SUM(CASE WHEN direction = '' THEN 1 ELSE 0 END)
        FROM traffic_by_vehicle
        """
    ).fetchone()
    print(f"null\t{null_count}")
    print(f"blank\t{blank_count}")
    connection.close()


if __name__ == "__main__":
    main()
