import argparse
import csv
import sqlite3
from pathlib import Path


DEFAULT_CSV = Path("02_Result/차종별_교통량_추출/전체교차로_260101~260430.csv")
DEFAULT_DB = Path("02_Result/차종별_교통량_추출/전체교차로_260101~260430.db")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="차종별 교차로 CSV를 SQLite DB로 변환합니다.")
    parser.add_argument("--csv", default=str(DEFAULT_CSV), help="입력 CSV 경로")
    parser.add_argument("--db", default=str(DEFAULT_DB), help="출력 SQLite DB 경로")
    parser.add_argument("--batch-size", type=int, default=50000, help="한 번에 INSERT할 행 수")
    return parser.parse_args()


def recreate_table(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        DROP TABLE IF EXISTS traffic_by_vehicle;

        CREATE TABLE traffic_by_vehicle (
            id INTEGER PRIMARY KEY,
            intersection TEXT NOT NULL,
            observed_at TEXT NOT NULL,
            approach_direction TEXT NOT NULL,
            direction TEXT NOT NULL,
            vehicle_type TEXT NOT NULL,
            traffic_volume INTEGER NOT NULL
        );
        """
    )


def create_indexes(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        CREATE INDEX IF NOT EXISTS idx_traffic_intersection_time
            ON traffic_by_vehicle (intersection, observed_at);
        CREATE INDEX IF NOT EXISTS idx_traffic_time
            ON traffic_by_vehicle (observed_at);
        CREATE INDEX IF NOT EXISTS idx_traffic_vehicle_type
            ON traffic_by_vehicle (vehicle_type);
        CREATE INDEX IF NOT EXISTS idx_traffic_direction
            ON traffic_by_vehicle (direction);
        """
    )


def to_record(row: list[str]) -> tuple[str, str, str, str, str, int]:
    return (
        row[0],
        row[1],
        row[2],
        row[3],
        row[4],
        int(row[5] or 0),
    )


def import_csv(csv_path: Path, db_path: Path, batch_size: int) -> int:
    db_path.parent.mkdir(parents=True, exist_ok=True)

    with sqlite3.connect(db_path) as conn:
        conn.execute("PRAGMA journal_mode = OFF")
        conn.execute("PRAGMA synchronous = OFF")
        conn.execute("PRAGMA temp_store = MEMORY")
        recreate_table(conn)

        insert_sql = """
            INSERT INTO traffic_by_vehicle (
                intersection,
                observed_at,
                approach_direction,
                direction,
                vehicle_type,
                traffic_volume
            )
            VALUES (?, ?, ?, ?, ?, ?)
        """

        total = 0
        batch: list[tuple[str, str, str, str, str, int]] = []

        with csv_path.open("r", newline="", encoding="utf-8-sig") as csv_file:
            reader = csv.reader(csv_file)
            header = next(reader, None)
            expected_header = ["교차로", "시간", "교차로 방향", "방향", "차종", "교통량"]
            if header != expected_header:
                raise ValueError(f"CSV 헤더가 예상과 다릅니다: {header}")

            for row in reader:
                if not row:
                    continue
                batch.append(to_record(row))
                if len(batch) >= batch_size:
                    conn.executemany(insert_sql, batch)
                    total += len(batch)
                    batch.clear()
                    print(f"imported {total:,} rows", flush=True)

            if batch:
                conn.executemany(insert_sql, batch)
                total += len(batch)
                print(f"imported {total:,} rows", flush=True)

        print("creating indexes...", flush=True)
        create_indexes(conn)
        conn.execute("VACUUM")

    return total


def main() -> None:
    args = parse_args()
    csv_path = Path(args.csv)
    db_path = Path(args.db)

    if not csv_path.exists():
        raise FileNotFoundError(csv_path)

    total = import_csv(csv_path, db_path, args.batch_size)
    print(f"done: {db_path} ({total:,} rows)")


if __name__ == "__main__":
    main()
