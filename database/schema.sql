PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS app_metadata (
  key TEXT PRIMARY KEY,
  value TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS import_sources (
  period TEXT PRIMARY KEY,
  source_origin TEXT NOT NULL,
  source_file TEXT NOT NULL DEFAULT '',
  imported_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS daily_movements (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  movement_date TEXT NOT NULL,
  period TEXT NOT NULL,
  office TEXT NOT NULL CHECK (office IN ('utara', 'selatan')),
  domestic_count INTEGER NOT NULL DEFAULT 0,
  foreign_count INTEGER NOT NULL DEFAULT 0,
  total_count INTEGER NOT NULL DEFAULT 0,
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
  UNIQUE (movement_date, office)
);

CREATE INDEX IF NOT EXISTS idx_daily_movements_period
  ON daily_movements (period);

CREATE TABLE IF NOT EXISTS agent_reports (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  period TEXT NOT NULL,
  category TEXT NOT NULL CHECK (category IN ('luar_negeri', 'dalam_negeri')),
  display_order INTEGER NOT NULL,
  row_no INTEGER,
  agent_name TEXT NOT NULL,
  movements INTEGER NOT NULL DEFAULT 0,
  estimated_revenue INTEGER NOT NULL DEFAULT 0,
  estimated_revenue_raw TEXT NOT NULL DEFAULT '0',
  is_total INTEGER NOT NULL DEFAULT 0 CHECK (is_total IN (0, 1)),
  source_title TEXT NOT NULL,
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_agent_reports_period_category
  ON agent_reports (period, category, display_order);

CREATE TABLE IF NOT EXISTS jetty_reports (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  period TEXT NOT NULL,
  category TEXT NOT NULL CHECK (category IN ('luar_negeri', 'dalam_negeri')),
  display_order INTEGER NOT NULL,
  row_no INTEGER,
  jetty_name TEXT NOT NULL,
  region_name TEXT NOT NULL DEFAULT '',
  movements INTEGER NOT NULL DEFAULT 0,
  is_total INTEGER NOT NULL DEFAULT 0 CHECK (is_total IN (0, 1)),
  source_title TEXT NOT NULL,
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_jetty_reports_period_category
  ON jetty_reports (period, category, display_order);

CREATE VIEW IF NOT EXISTS v_monthly_office_totals AS
SELECT
  period,
  office,
  SUM(domestic_count) AS domestic_total,
  SUM(foreign_count) AS foreign_total,
  SUM(total_count) AS movement_total
FROM daily_movements
GROUP BY period, office;

CREATE VIEW IF NOT EXISTS v_monthly_agent_totals AS
SELECT
  period,
  category,
  SUM(movements) AS movement_total,
  SUM(estimated_revenue) AS estimated_revenue_total
FROM agent_reports
WHERE is_total = 0
GROUP BY period, category;

CREATE VIEW IF NOT EXISTS v_monthly_jetty_totals AS
SELECT
  period,
  category,
  SUM(movements) AS movement_total
FROM jetty_reports
WHERE is_total = 0
GROUP BY period, category;
