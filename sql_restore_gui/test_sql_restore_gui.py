import os
import sys
import tempfile
import tkinter as tk
import unittest
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import sql_restore_gui


class TestSqlRestoreGui(unittest.TestCase):
    def test_parse_backup_filename_single_and_striped(self):
        # Striped .bak produced by sql_backup_gui
        db, ts, idx, total = sql_restore_gui.parse_backup_filename(
            "EnterpriseWarehouseDB_20261021_100000_stripe03of08.bak"
        )
        self.assertEqual(db, "EnterpriseWarehouseDB")
        self.assertEqual(ts, "20261021_100000")
        self.assertEqual(idx, 3)
        self.assertEqual(total, 8)

        # Single .bak produced by sql_backup_gui
        db2, ts2, idx2, total2 = sql_restore_gui.parse_backup_filename(
            "SalesDB_20261021_100000.bak"
        )
        self.assertEqual(db2, "SalesDB")
        self.assertEqual(ts2, "20261021_100000")
        self.assertEqual(idx2, 1)
        self.assertEqual(total2, 1)

        # Generic _partXofY naming without timestamp
        db3, ts3, idx3, total3 = sql_restore_gui.parse_backup_filename(
            "BillingDB_part2of4.bak"
        )
        self.assertEqual(db3, "BillingDB")
        self.assertEqual(ts3, "")
        self.assertEqual(idx3, 2)
        self.assertEqual(total3, 4)

    def test_scan_backup_source_folder_groups_stripes_and_zips_and_filters_latest(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            # 1. Create older 2-stripe backup set for WarehouseDB
            wh_dir = os.path.join(tmpdir, "WarehouseDB")
            os.makedirs(wh_dir, exist_ok=True)
            for i in (1, 2):
                p = os.path.join(wh_dir, f"WarehouseDB_20261020_090000_stripe0{i}of02.bak")
                Path(p).write_bytes(b"OLD_STRIPE_" * 100)

            # 2. Create newer 4-stripe backup set for WarehouseDB
            for i in range(1, 5):
                p = os.path.join(wh_dir, f"WarehouseDB_20261021_110000_stripe0{i}of04.bak")
                Path(p).write_bytes(b"NEW_STRIPE_DATA_" * 200)

            # 3. Create a .zip archive containing 2 striped .bak files for OrdersDB
            zip_dir = os.path.join(tmpdir, "Zipped")
            os.makedirs(zip_dir, exist_ok=True)
            zip_path = os.path.join(zip_dir, "OrdersDB_20261021_120000.zip")
            with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=1) as zf:
                zf.writestr("OrdersDB_20261021_120000_stripe02of02.bak", b"ORDERS_S2_" * 150)
                zf.writestr("OrdersDB_20261021_120000_stripe01of02.bak", b"ORDERS_S1_" * 150)

            # 4. Ensure _RestoreStaging folder is ignored during scan
            staging_dir = os.path.join(tmpdir, "_RestoreStaging", "IgnoredDB")
            os.makedirs(staging_dir, exist_ok=True)
            Path(os.path.join(staging_dir, "IgnoredDB_20261021_120000.bak")).write_bytes(b"IGNORE")

            # Scan with latest_only_per_db=True
            latest_sets = sql_restore_gui.scan_backup_source_folder(tmpdir, latest_only_per_db=True)
            self.assertEqual(len(latest_sets), 2)
            by_name = {s.source_db_name: s for s in latest_sets}
            self.assertIn("WarehouseDB", by_name)
            self.assertIn("OrdersDB", by_name)

            wh_set = by_name["WarehouseDB"]
            self.assertEqual(wh_set.timestamp_label, "20261021_110000")
            self.assertEqual(wh_set.expected_stripes, 4)
            self.assertEqual(wh_set.found_stripes, 4)
            self.assertTrue(wh_set.is_complete)
            self.assertTrue(wh_set.stripe_paths[0].endswith("stripe01of04.bak"))
            self.assertTrue(wh_set.stripe_paths[3].endswith("stripe04of04.bak"))

            ord_set = by_name["OrdersDB"]
            self.assertEqual(ord_set.source_type, "zip")
            self.assertEqual(ord_set.expected_stripes, 2)
            self.assertEqual(ord_set.found_stripes, 2)
            self.assertEqual(
                ord_set.stripe_paths,
                [
                    "OrdersDB_20261021_120000_stripe01of02.bak",
                    "OrdersDB_20261021_120000_stripe02of02.bak",
                ],
            )

            # Scan with latest_only_per_db=False should return all 3 sets (including older WarehouseDB)
            all_sets = sql_restore_gui.scan_backup_source_folder(tmpdir, latest_only_per_db=False)
            self.assertEqual(len(all_sets), 3)

    def test_fast_unzip_backup_set_and_staging_cleanup(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            zip_path = os.path.join(tmpdir, "CustomersDB_20261021_100000.zip")
            staging_root = os.path.join(tmpdir, "_RestoreStaging")
            with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=1) as zf:
                zf.writestr("CustomersDB_20261021_100000_stripe01of02.bak", b"STRIPE_1_PAYLOAD" * 200)
                zf.writestr("CustomersDB_20261021_100000_stripe02of02.bak", b"STRIPE_2_PAYLOAD" * 200)

            extract_dir, paths, elapsed = sql_restore_gui.fast_unzip_backup_set(
                zip_path=zip_path,
                staging_root=staging_root,
                db_name="CustomersDB",
                timestamp_label="20261021_100000",
                prefer_7zip=False,
            )
            self.assertEqual(len(paths), 2)
            self.assertTrue(os.path.isfile(paths[0]))
            self.assertTrue(os.path.isfile(paths[1]))
            self.assertGreaterEqual(elapsed, 0.0)

            cleaned = sql_restore_gui.cleanup_extracted_staging_files(extract_dir, paths)
            self.assertTrue(cleaned)
            for p in paths:
                self.assertFalse(os.path.exists(p))

    def test_build_move_clauses_and_renamed_target_db(self):
        logical_files = [
            sql_restore_gui.LogicalFileInfo(
                logical_name="SalesDB_Data",
                physical_name=r"D:\OldData\SalesDB.mdf",
                file_type="D",
                file_id=1,
            ),
            sql_restore_gui.LogicalFileInfo(
                logical_name="SalesDB_Secondary",
                physical_name=r"D:\OldData\SalesDB_sec.ndf",
                file_type="D",
                file_id=2,
            ),
            sql_restore_gui.LogicalFileInfo(
                logical_name="SalesDB_Log",
                physical_name=r"E:\OldLogs\SalesDB_log.ldf",
                file_type="L",
                file_id=3,
            ),
        ]
        # Side-by-side restore as SalesDB_UAT
        moves = sql_restore_gui.build_move_clauses(
            source_db_name="SalesDB",
            target_db_name="SalesDB_UAT",
            logical_files=logical_files,
            target_data_folder=r"F:\SqlData",
            target_log_folder=r"G:\SqlLogs",
        )
        self.assertEqual(
            moves,
            [
                ("SalesDB_Data", r"F:\SqlData\SalesDB_UAT.mdf"),
                ("SalesDB_Secondary", r"F:\SqlData\SalesDB_UAT_sec.ndf"),
                ("SalesDB_Log", r"G:\SqlLogs\SalesDB_UAT_log.ldf"),
            ],
        )

    def test_build_restore_tsql_high_speed_striped(self):
        stripes = [
            r"D:\Backups\AnalyticsDB\AnalyticsDB_20261021_100000_stripe01of04.bak",
            r"D:\Backups\AnalyticsDB\AnalyticsDB_20261021_100000_stripe02of04.bak",
            r"D:\Backups\AnalyticsDB\AnalyticsDB_20261021_100000_stripe03of04.bak",
            r"D:\Backups\AnalyticsDB\AnalyticsDB_20261021_100000_stripe04of04.bak",
        ]
        tsql = sql_restore_gui.build_restore_tsql(
            source_db_name="AnalyticsDB",
            target_db_name="AnalyticsDB",
            stripe_file_paths=stripes,
            relocate_files=True,
            target_data_folder=r"E:\SqlData",
            target_log_folder=r"F:\SqlLogs",
            with_replace=True,
            kill_connections=True,
            recovery_mode="RECOVERY",
            high_speed_tuning=True,
            custom_buffer_count=0,
            keep_cdc=True,
            stats_interval=5,
        )
        self.assertIn("SET SINGLE_USER WITH ROLLBACK IMMEDIATE;", tsql)
        self.assertIn("RESTORE DATABASE [AnalyticsDB]", tsql)
        self.assertIn("stripe01of04.bak", tsql)
        self.assertIn("stripe04of04.bak", tsql)
        self.assertIn("MOVE N'AnalyticsDB' TO N'E:\\SqlData\\AnalyticsDB.mdf'", tsql)
        self.assertIn("MOVE N'AnalyticsDB_log' TO N'F:\\SqlLogs\\AnalyticsDB_log.ldf'", tsql)
        self.assertIn("REPLACE", tsql)
        self.assertIn("KEEP_CDC", tsql)
        self.assertIn("RECOVERY", tsql)
        self.assertIn("MAXTRANSFERSIZE = 4194304", tsql)
        self.assertIn("BUFFERCOUNT = 32", tsql)  # 4 stripes * 8 = 32
        self.assertIn("SET MULTI_USER;", tsql)

    def test_sort_backup_sets_largest_first_lpt(self):
        s_small = sql_restore_gui.DiscoveredBackupSet(
            set_id="1",
            source_db_name="SmallDB",
            timestamp_label="20261021",
            source_type="bak",
            total_size_bytes=2 * 1024**3,
        )
        s_huge = sql_restore_gui.DiscoveredBackupSet(
            set_id="2",
            source_db_name="HugeDB",
            timestamp_label="20261021",
            source_type="bak",
            total_size_bytes=80 * 1024**3,
        )
        s_med = sql_restore_gui.DiscoveredBackupSet(
            set_id="3",
            source_db_name="MediumDB",
            timestamp_label="20261021",
            source_type="bak",
            total_size_bytes=18 * 1024**3,
        )
        ordered = sql_restore_gui.sort_backup_sets_largest_first([s_small, s_huge, s_med])
        self.assertEqual([s.source_db_name for s in ordered], ["HugeDB", "MediumDB", "SmallDB"])

    def test_gui_instantiation_and_preview(self):
        root = tk.Tk()
        root.withdraw()
        try:
            app = sql_restore_gui.SqlRestoreApp(root)
            app.all_backup_sets = [
                sql_restore_gui.DiscoveredBackupSet(
                    set_id="bak::1",
                    source_db_name="HugeAnalyticsDB",
                    timestamp_label="20261021_120000",
                    source_type="bak",
                    stripe_paths=[
                        r"C:\SqlBackups\HugeAnalyticsDB\HugeAnalyticsDB_20261021_120000_stripe01of04.bak",
                        r"C:\SqlBackups\HugeAnalyticsDB\HugeAnalyticsDB_20261021_120000_stripe02of04.bak",
                        r"C:\SqlBackups\HugeAnalyticsDB\HugeAnalyticsDB_20261021_120000_stripe03of04.bak",
                        r"C:\SqlBackups\HugeAnalyticsDB\HugeAnalyticsDB_20261021_120000_stripe04of04.bak",
                    ],
                    expected_stripes=4,
                    found_stripes=4,
                    total_size_bytes=32 * 1024**3,
                )
            ]
            app.selected_set_ids = {"bak::1"}
            app.var_db_suffix.set("_RESTORED")
            app.refresh_backup_tree()
            app.update_tsql_preview()

            preview_text = app.txt_tsql.get("1.0", tk.END)
            self.assertIn("RESTORE DATABASE [HugeAnalyticsDB_RESTORED]", preview_text)
            self.assertIn("stripe01of04.bak", preview_text)
            self.assertIn("stripe04of04.bak", preview_text)
            self.assertIn("MAXTRANSFERSIZE = 4194304", preview_text)
            self.assertIn("BUFFERCOUNT = 32", preview_text)
        finally:
            root.destroy()

    def test_powershell_edition_helpers_and_syntax(self):
        import subprocess

        ps1_path = Path(__file__).resolve().parent / "sql_restore_gui.ps1"
        self.assertTrue(ps1_path.is_file())

        with tempfile.TemporaryDirectory() as tmpdir:
            db_dir = os.path.join(tmpdir, "OrdersDB")
            os.makedirs(db_dir, exist_ok=True)
            for i in (1, 2):
                Path(os.path.join(db_dir, f"OrdersDB_20261021_100000_stripe0{i}of02.bak")).write_bytes(
                    b"MOCK_BAK_" * 50
                )

            ps_cmd = (
                f". '{ps1_path}' -NoGui; "
                f"$sets = @(Find-BackupRestoreSets -SourceFolder '{tmpdir}' -LatestOnlyPerDb $true); "
                f"if ($sets.Count -ne 1 -or $sets[0].SourceDbName -ne 'OrdersDB' -or $sets[0].FoundStripes -ne 2) {{ throw 'Find-BackupRestoreSets failed' }}; "
                f"$sql = Get-RestoreDatabaseTsql -SourceDbName 'OrdersDB' -TargetDbName 'OrdersDB_UAT' -StripeFilePaths $sets[0].StripePaths -TargetDataFolder 'E:\\Data' -TargetLogFolder 'F:\\Logs'; "
                f"if ($sql -notmatch 'MAXTRANSFERSIZE = 4194304') {{ throw 'T-SQL check failed' }}; "
                f"Write-Output 'PS1_OK'"
            )
            proc = subprocess.run(
                [
                    "powershell.exe",
                    "-NoProfile",
                    "-NonInteractive",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    ps_cmd,
                ],
                capture_output=True,
                text=True,
                timeout=20,
            )
            self.assertEqual(proc.returncode, 0, msg=proc.stderr or proc.stdout)
            self.assertIn("PS1_OK", proc.stdout)


if __name__ == "__main__":
    unittest.main()
