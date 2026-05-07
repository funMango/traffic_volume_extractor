import oracledb, os
from dotenv import load_dotenv
from pathlib import Path

load_dotenv(Path(__file__).parent.parent / "00_Data" / ".env")

print(f"HOST: {os.getenv('DB_HOST')}")
print(f"PORT: {os.getenv('DB_PORT')}")
print(f"SERVICE: {os.getenv('DB_SERVICE')}")
print(f"USER: {os.getenv('DB_USER')}")
print("연결 시도 중...")

try:
    conn = oracledb.connect(
        user=os.getenv("DB_USER").strip(),
        password=os.getenv("DB_PASSWORD").strip(),
        host=os.getenv("DB_HOST").strip(),
        port=int(os.getenv("DB_PORT", "1521").strip()),
        service_name=os.getenv("DB_SERVICE").strip(),
        # tcp_connect_timeout=5,  # 원본 connect_db()와 동일하게 timeout 없음
    )
    print("연결 성공!")
    with conn.cursor() as cur:
        print("쿼리 실행 중...")
        cur.execute("SELECT NODE_ID, CRSRD_NM FROM M_CRSRD_INF ORDER BY CRSRD_NM")
        rows = cur.fetchall()
        print(f"교차로 수: {len(rows)}개")
    conn.close()
except Exception as e:
    print(f"오류: {e}")
