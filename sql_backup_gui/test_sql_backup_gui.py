import os
import sys
import tempfile
import tkinter as tk
import unittest
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import sql_backup_gui


class TestSqlBackupGui(unittest.TestCase):
    def test_calculate_stream_count_threshold_8gb(self):
        # Databases <= 8 GB should always use 1 stream
        self.assertEqual(sql_backup_gui.calculate_stream_count(0.25, threshold_gb=8.0, max_streams=8), 1)
        self.assertEqual(sql_backup_gui.calculate_stream_count(8.0, threshold_gb=8.0, max_streams=8), 1)

        # Databases > 8 GB should split dynamically into parallel striped streams
        self.assertEqual(sql_backup_gui.calculate_stream_count(8.01, threshold_gb=8.0, max_streams=8), 2)
        self.assertEqual(sql_backup_gui.calculate_stream_count(16.0, threshold_gb=8.0, max_streams=8), 2)
        self.assertEqual(sql_backup_gui.calculate_stream_count(16.5, threshold_gb=8.0, max_streams=8), 3)
        self.assertEqual(sql_backup_gui.calculate_stream_count(28.0, threshold_gb=8.0, max_streams=8), 4)
        self.assertEqual(sql_backup_gui.calculate_stream_count(95.0, threshold_gb=8.0, max_streams=8), 8)

    def test_build_backup_file_paths_single_and_striped(self):
        # Single stream (<= 8 GB)
        subfolder1, paths1 = sql_backup_gui.build_backup_file_paths(
            target_folder=r"D:\SqlBackups",
            db_name="SalesDB",
            stream_count=1,
            timestamp="20261021_100000",
        )
        self.assertEqual(subfolder1, r"D:\SqlBackups\SalesDB")
        self.assertEqual(paths1, [r"D:\SqlBackups\SalesDB\SalesDB_20261021_100000.bak"])

        # Multi-stream (> 8 GB -> 4 streams)
        subfolder4, paths4 = sql_backup_gui.build_backup_file_paths(
            target_folder=r"D:\SqlBackups\\",
            db_name="WarehouseDB",
            stream_count=4,
            timestamp="20261021_100000",
        )
        self.assertEqual(subfolder4, r"D:\SqlBackups\WarehouseDB")
        self.assertEqual(len(paths4), 4)
        self.assertEqual(paths4[0], r"D:\SqlBackups\WarehouseDB\WarehouseDB_20261021_100000_stripe01of04.bak")
        self.assertEqual(paths4[3], r"D:\SqlBackups\WarehouseDB\WarehouseDB_20261021_100000_stripe04of04.bak")

    def test_build_zip_file_path_and_fast_zip_execution(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            target_dir = os.path.join(tmpdir, "Backups")
            zip_dir, zip_path = sql_backup_gui.build_zip_file_path(
                target_folder=target_dir,
                db_name="OrdersDB",
                timestamp="20261021_110000",
            )
            self.assertTrue(zip_dir.endswith("Zipped"))
            self.assertTrue(zip_path.endswith("OrdersDB_20261021_110000.zip"))

            # Create two mock striped .bak files and fast-zip them (using stdlib compresslevel=1)
            subfolder, bak_paths = sql_backup_gui.build_backup_file_paths(
                target_folder=target_dir,
                db_name="OrdersDB",
                stream_count=2,
                timestamp="20261021_110000",
            )
            os.makedirs(subfolder, exist_ok=True)
            for idx, p in enumerate(bak_paths, start=1):
                Path(p).write_bytes((f"MOCK_SQL_STRIPE_{idx}_" * 500).encode("utf-8"))

            out_zip, elapsed = sql_backup_gui.fast_zip_backup_files(
                file_paths=bak_paths,
                zip_path=zip_path,
                delete_source_after_zip=True,
                source_subfolder=subfolder,
                prefer_7zip=False,
            )
            self.assertEqual(out_zip, zip_path)
            self.assertTrue(os.path.isfile(zip_path))
            self.assertGreaterEqual(elapsed, 0.0)

            # Verify archive entries
            with zipfile.ZipFile(zip_path, "r") as zf:
                names = sorted(zf.namelist())
                self.assertEqual(
                    names,
                    [
                        "OrdersDB_20261021_110000_stripe01of02.bak",
                        "OrdersDB_20261021_110000_stripe02of02.bak",
                    ],
                )
            # Verify raw .bak files were cleaned up when delete_source_after_zip=True
            for p in bak_paths:
                self.assertFalse(os.path.exists(p))

    def test_batch_folder_zip_and_delete_unzipped_backups(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            target_dir = os.path.join(tmpdir, "SqlBackups")
            zip_dir, batch_zip_path, parent_name = sql_backup_gui.build_batch_zip_file_path(
                target_folder=target_dir,
                timestamp="20261021_223000",
            )
            self.assertEqual(parent_name, "SqlBackups")
            self.assertTrue(batch_zip_path.endswith("SqlBackups_20261021_223000.zip"))

            # Create two DB subfolders with mock .bak files
            sub1, paths1 = sql_backup_gui.build_backup_file_paths(
                target_folder=target_dir,
                db_name="AlphaDB",
                stream_count=1,
                timestamp="20261021_223000",
            )
            sub2, paths2 = sql_backup_gui.build_backup_file_paths(
                target_folder=target_dir,
                db_name="BetaWarehouseDB",
                stream_count=2,
                timestamp="20261021_223000",
            )
            os.makedirs(sub1, exist_ok=True)
            os.makedirs(sub2, exist_ok=True)
            for p in paths1 + paths2:
                Path(p).write_bytes(b"MOCK_BACKUP_DATA_" * 300)

            out_zip, elapsed, count = sql_backup_gui.fast_zip_target_folder_batch(
                target_folder=target_dir,
                zip_path=batch_zip_path,
                db_subfolders=[sub1, sub2],
                delete_source_after_zip=True,
            )
            self.assertEqual(out_zip, batch_zip_path)
            self.assertEqual(count, 3)
            self.assertTrue(os.path.isfile(batch_zip_path))

            # Verify all DB folders are inside the zipped parent folder "SqlBackups/"
            with zipfile.ZipFile(batch_zip_path, "r") as zf:
                names = sorted(zf.namelist())
                self.assertEqual(
                    names,
                    [
                        "SqlBackups/AlphaDB/AlphaDB_20261021_223000.bak",
                        "SqlBackups/BetaWarehouseDB/BetaWarehouseDB_20261021_223000_stripe01of02.bak",
                        "SqlBackups/BetaWarehouseDB/BetaWarehouseDB_20261021_223000_stripe02of02.bak",
                    ],
                )

            # Verify unzipped DB backup folders were deleted after completion
            self.assertFalse(os.path.exists(sub1))
            self.assertFalse(os.path.exists(sub2))

    def test_create_subdir_tsql_wraps_error_183_in_try_catch(self):
        sql = sql_backup_gui.build_create_subdir_tsql(r"C:\Users\cdobbs4\SqlBackups\TSTSystemCatalog")
        self.assertIn("BEGIN TRY", sql)
        self.assertIn("EXEC master.dbo.xp_create_subdir N'C:\\Users\\cdobbs4\\SqlBackups\\TSTSystemCatalog'", sql)
        self.assertIn("END CATCH;", sql)

        with tempfile.TemporaryDirectory() as tmpdir:
            sub = os.path.join(tmpdir, "TSTSystemCatalog")
            sql_backup_gui.ensure_local_backup_dir_permissions(tmpdir, sub)
            self.assertTrue(os.path.isdir(sub))

    def test_build_backup_tsql_copy_only_and_stripes(self):
        _, paths = sql_backup_gui.build_backup_file_paths(
            target_folder=r"E:\Backups",
            db_name="Reporting[Prod]",
            stream_count=3,
            timestamp="20261021_120000",
        )
        tsql = sql_backup_gui.build_backup_tsql(
            db_name="Reporting[Prod]",
            file_paths=paths,
            copy_only=True,
            compression=True,
            checksum=True,
            stats_interval=5,
        )
        self.assertIn("BACKUP DATABASE [Reporting[Prod]]]", tsql)
        self.assertIn("DISK = N'E:\\Backups\\Reporting[Prod]\\Reporting[Prod]_20261021_120000_stripe01of03.bak'", tsql)
        self.assertIn("DISK = N'E:\\Backups\\Reporting[Prod]\\Reporting[Prod]_20261021_120000_stripe02of03.bak'", tsql)
        self.assertIn("DISK = N'E:\\Backups\\Reporting[Prod]\\Reporting[Prod]_20261021_120000_stripe03of03.bak'", tsql)
        self.assertIn("COPY_ONLY", tsql)
        self.assertIn("COMPRESSION", tsql)
        self.assertIn("CHECKSUM", tsql)
        self.assertIn("STATS = 5", tsql)

    def test_connection_string_generation(self):
        win_cfg = sql_backup_gui.SqlConnectionConfig(server=r"SQLPROD01\INST1", auth_mode="windows")
        conn_win = win_cfg.to_dotnet_connection_string()
        self.assertIn(r"Server=SQLPROD01\INST1", conn_win)
        self.assertIn("Integrated Security=True", conn_win)

        sql_cfg = sql_backup_gui.SqlConnectionConfig(
            server="10.0.0.25,1433",
            auth_mode="sql",
            username="backup_admin",
            password="My;SecretPassword",
        )
        conn_sql = sql_cfg.to_dotnet_connection_string()
        self.assertIn("Integrated Security=False", conn_sql)
        self.assertIn("User ID=backup_admin", conn_sql)
        self.assertIn("Password=My;;SecretPassword", conn_sql)

    def test_gui_instantiation_and_preview(self):
        root = tk.Tk()
        root.withdraw()
        try:
            app = sql_backup_gui.SqlBackupApp(root)
            app.all_databases = [
                sql_backup_gui.DatabaseInfo(name="SmallAppDB", size_mb=2048.0),      # 2 GB -> 1 stream
                sql_backup_gui.DatabaseInfo(name="HugeAnalyticsDB", size_mb=25600.0),  # 25 GB -> 4 streams
            ]
            app.selected_db_names = {"SmallAppDB", "HugeAnalyticsDB"}
            app.var_target_folder.set(r"C:\SqlBackups")
            app.var_zip_after_backup.set(True)
            app.refresh_database_tree()
            app.update_tsql_preview()

            preview_text = app.txt_tsql.get("1.0", tk.END)
            self.assertIn("BACKUP DATABASE [SmallAppDB]", preview_text)
            self.assertIn("BACKUP DATABASE [HugeAnalyticsDB]", preview_text)
            self.assertIn("stripe01of04.bak", preview_text)
            self.assertIn("stripe04of04.bak", preview_text)
            self.assertIn("COPY_ONLY", preview_text)
            self.assertIn("Fast-Zip Whole Target Folder", preview_text)

            app.var_zip_mode.set("per_db")
            app.update_tsql_preview()
            self.assertIn("Fast-Zip Output", app.txt_tsql.get("1.0", tk.END))
        finally:
            root.destroy()


if __name__ == "__main__":
    unittest.main()
