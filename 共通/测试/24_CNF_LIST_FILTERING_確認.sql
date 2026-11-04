-- ============================================================
-- 24 CNF_LIST_FILTERING_TBL（リスト切替フィルタリング設定）
--   実機DB 確認用
--   2026/10/29  孫
--   阿部さんからの確認依頼（実機DBの FILTERING_DEF_ID は 7 か？）への回答用
-- ============================================================

-- 接続
--   psql -h amldb01 -U aml_user -d amldb


-- ------------------------------------------------------------
-- ① そもそもテーブルがあるか・列は何か
-- ------------------------------------------------------------
\dt aml.cnf_list*
\d  aml.cnf_list_filtering_tbl

--   見つからない場合は検索パスが違う可能性がある。
--     \dt *.*filtering*
--   で全スキーマを探すこと。


-- ------------------------------------------------------------
-- ② 現状の中身（これが本命。結果をそのまま阿部さんに貼る）
-- ------------------------------------------------------------
SELECT filtering_list_id,
       filtering_def_id
  FROM aml.cnf_list_filtering_tbl
 ORDER BY filtering_list_id, filtering_def_id;

--   期待（マスタ定義書 2026/10/28 修正後の姿）
--     filtering_list_id = 0 に対して
--     filtering_def_id  = 1,2,3,4,5,7,8,9,10,11,12,13,14,15,16   … 全15行
--
--   判定ポイント
--     6  がある  → 26 と同じ不具合。全銀協（個人）ではなく警察庁（IB）を見ている
--     7  が無い  → 同上
--     16 が無い  → L&F 追加分（BファイルC情報）が実機に未反映


-- ------------------------------------------------------------
-- ③ ②を1行にまとめた版（メールに貼りやすい）
-- ------------------------------------------------------------
SELECT filtering_list_id,
       string_agg(filtering_def_id::text, ',' ORDER BY filtering_def_id) AS def_ids,
       count(*) AS cnt
  FROM aml.cnf_list_filtering_tbl
 GROUP BY filtering_list_id
 ORDER BY filtering_list_id;


-- ------------------------------------------------------------
-- ④ 6 / 7 / 16 だけ数える（一目で判定したいとき）
-- ------------------------------------------------------------
SELECT count(*) FILTER (WHERE filtering_def_id = 6)  AS def6,
       count(*) FILTER (WHERE filtering_def_id = 7)  AS def7,
       count(*) FILTER (WHERE filtering_def_id = 16) AS def16,
       count(*)                                      AS total
  FROM aml.cnf_list_filtering_tbl;

--   正  def6 = 0 / def7 = 1 / def16 = 1 / total = 15


-- ------------------------------------------------------------
-- ⑤ 参考：DEF_ID が何のリストを指しているかを突き合わせる
-- ------------------------------------------------------------
SELECT d.filtering_def_id,
       d.filtering_def_name,
       o.origin,
       o.origin_jp
  FROM mst_filtering_def_tbl d
  LEFT JOIN mst_lstorigin_tbl o ON o.filtering_def_id = d.filtering_def_id
 ORDER BY d.filtering_def_id;

--   列名が違ったら ①の \d mst_filtering_def_tbl で確認して直すこと。
--   6 = 警察庁（IB）NPI ／ 7 = 全銀協（個人）ASI のはず。


-- ============================================================
-- ここから下は【実行しないでください】
--
-- ②〜④の結果が期待と違った場合の修正案（下書き）。
-- 2026/10/28 に阿部さんからいただいた了承は 26 だけが対象で、
-- 24 は含まれていない。必ず別途了承を得てから実行すること。
-- ============================================================

-- -- バックアップ
-- CREATE TABLE aml.cnf_list_filtering_tbl_bk20261029 AS
--   SELECT * FROM aml.cnf_list_filtering_tbl;
--
-- BEGIN;
--
-- -- 6（警察庁（IB））→ 7（全銀協（個人））　1件のはず
-- UPDATE aml.cnf_list_filtering_tbl
--    SET filtering_def_id = 7
--  WHERE filtering_list_id = 0
--    AND filtering_def_id  = 6;
--
-- -- L&F 追加分（FILTERING_DEF_ID = 16）　1件
-- INSERT INTO aml.cnf_list_filtering_tbl (filtering_list_id, filtering_def_id)
--   VALUES (0, 16);
--   -- ※ ①の \d で列が3つ以上あった場合は列名と値を合わせること
--
-- -- UPDATE 1件 / INSERT 1件 を目で確認してから
-- COMMIT;     -- 違っていたら ROLLBACK;
--
-- -- 反映確認
-- -- ③ を再実行して 1,2,3,4,5,7,8,9,10,11,12,13,14,15,16 になること
