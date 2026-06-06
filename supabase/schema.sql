-- ============================================================
-- 시약 재고 관리 시스템 - Supabase 스키마
-- Supabase 대시보드 > SQL Editor 에서 실행하세요
-- ============================================================

-- 1. 시약 테이블
CREATE TABLE IF NOT EXISTS reagents (
  id              INTEGER PRIMARY KEY,
  name            TEXT    NOT NULL DEFAULT '',
  manufacturer    TEXT    DEFAULT '',
  lot_no          TEXT    DEFAULT '',
  received_qty    INTEGER DEFAULT 0,
  current_stock   INTEGER DEFAULT 0,
  min_stock       INTEGER DEFAULT 0,
  expiry_date     TEXT    DEFAULT '',
  total_dispatched INTEGER DEFAULT 0,
  updated_at      TIMESTAMPTZ DEFAULT NOW()
);

-- 2. 출고 이력 테이블
CREATE TABLE IF NOT EXISTS dispatch_logs (
  id            BIGINT  PRIMARY KEY,
  reagent_id    INTEGER,
  reagent_name  TEXT    NOT NULL DEFAULT '',
  lot_no        TEXT    DEFAULT '',
  qty           INTEGER DEFAULT 1,
  datetime      TEXT    NOT NULL DEFAULT '',
  created_at    TIMESTAMPTZ DEFAULT NOW()
);

-- 3. RLS 활성화
ALTER TABLE reagents      ENABLE ROW LEVEL SECURITY;
ALTER TABLE dispatch_logs ENABLE ROW LEVEL SECURITY;

-- 4. 정책: anon 키로 전체 허용 (내부망 전용 도구)
DROP POLICY IF EXISTS "anon_all_reagents"      ON reagents;
DROP POLICY IF EXISTS "anon_all_dispatch_logs" ON dispatch_logs;

CREATE POLICY "anon_all_reagents"
  ON reagents FOR ALL TO anon USING (true) WITH CHECK (true);

CREATE POLICY "anon_all_dispatch_logs"
  ON dispatch_logs FOR ALL TO anon USING (true) WITH CHECK (true);

-- 5. 신규 컬럼 추가 (이미 있으면 무시)
ALTER TABLE reagents ADD COLUMN IF NOT EXISTS created_at TIMESTAMPTZ DEFAULT NOW();
ALTER TABLE reagents ADD COLUMN IF NOT EXISTS reagent_type TEXT DEFAULT 'Reagent';

-- 6. 입고 이력 테이블
CREATE TABLE IF NOT EXISTS inbound_logs (
  id           BIGINT  PRIMARY KEY,
  reagent_id   INTEGER,
  reagent_name TEXT    NOT NULL DEFAULT '',
  lot_no       TEXT    DEFAULT '',
  qty          INTEGER DEFAULT 1,
  datetime     TEXT    NOT NULL DEFAULT '',
  notes        TEXT    DEFAULT '',
  created_at   TIMESTAMPTZ DEFAULT NOW()
);
ALTER TABLE inbound_logs ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "anon_all_inbound_logs" ON inbound_logs;
CREATE POLICY "anon_all_inbound_logs"
  ON inbound_logs FOR ALL TO anon USING (true) WITH CHECK (true);

-- 7. 초기 시약 데이터 (이미 있으면 건너뜀)
INSERT INTO reagents (id, name, manufacturer, lot_no, received_qty, current_stock, min_stock, expiry_date, total_dispatched) VALUES
  (1, '혈액형 검사 시약 (ABO/Rh)', 'Bio-Rad',  'BR2024-001', 100,  8, 10, '2026-06-15',  92),
  (2, 'CBC 희석액',                 'Sysmex',   'SX2024-102', 200, 45, 30, '2026-08-20', 155),
  (3, 'PT/APTT 응고 시약',          'Stago',    'ST2024-055',  50, 12, 15, '2026-06-10',  38),
  (4, 'HbA1c 측정 시약',            'Tosoh',    'TS2024-210',  80, 30, 20, '2026-09-05',  50),
  (5, 'Troponin I 키트',            'Abbott',   'AB2024-334',  60,  5,  8, '2026-07-30',  55),
  (6, '요검사 스트립',               'Roche',    'RC2024-412', 300, 90, 50, '2026-10-15', 210),
  (7, 'CRP 정량 시약',               'Beckman',  'BK2024-078',  40, 18, 10, '2026-06-18',  22)
ON CONFLICT (id) DO NOTHING;
