import sys
import tempfile
import tkinter as tk
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent))
import aws_cp_gui


class TestAwsCpGui(unittest.TestCase):
    def test_parse_and_build_s3_uri(self):
        b, k = aws_cp_gui.parse_s3_uri("s3://my-bucket/folder/file.txt")
        self.assertEqual(b, "my-bucket")
        self.assertEqual(k, "folder/file.txt")

        b2, k2 = aws_cp_gui.parse_s3_uri("s3://my-bucket/")
        self.assertEqual(b2, "my-bucket")
        self.assertEqual(k2, "")

        self.assertEqual(aws_cp_gui.build_s3_uri("my-bucket", "a/b.txt"), "s3://my-bucket/a/b.txt")
        self.assertEqual(aws_cp_gui.build_s3_uri("my-bucket", ""), "s3://my-bucket/")

    def test_build_aws_cp_command_args(self):
        args = aws_cp_gui.build_aws_cp_command_args(
            source=r"C:\data\report 1.csv",
            target="s3://corp-bucket/reports/",
            region="eu-west-2",
            output_format="json",
            recursive=False,
            dryrun=True,
            only_show_errors=True,
            storage_class="STANDARD_IA",
            exclude_pattern="*.tmp",
            include_pattern="*.csv",
        )
        preview = aws_cp_gui.format_command_preview(args)
        self.assertIn('aws s3 cp "C:\\data\\report 1.csv" s3://corp-bucket/reports/', preview)
        self.assertIn("--dryrun", preview)
        self.assertIn("--only-show-errors", preview)
        self.assertIn("--storage-class STANDARD_IA", preview)
        self.assertIn("--region eu-west-2", preview)
        self.assertIn("--output json", preview)

    def test_sigv4_signed_request_generation(self):
        creds = aws_cp_gui.AWSCredentialsConfig(
            access_key_id="AKIAIOSFODNN7EXAMPLE",
            secret_access_key="wJalrXUtnFEMI/K7MDENG/bPxRfiCYEXAMPLEKEY",
            region="us-east-1",
            output_format="json",
        )
        client = aws_cp_gui.StdlibS3Client(creds)
        req = client._build_signed_request(
            method="GET",
            bucket="example-bucket",
            key="folder/my file (1).txt",
            query_params={"list-type": "2", "prefix": "folder/"},
        )
        self.assertTrue(req.full_url.startswith("https://example-bucket.s3.us-east-1.amazonaws.com/"))
        auth = req.headers.get("Authorization")
        self.assertIsNotNone(auth)
        self.assertIn("AWS4-HMAC-SHA256 Credential=AKIAIOSFODNN7EXAMPLE/", auth)
        self.assertIn("SignedHeaders=host;x-amz-content-sha256;x-amz-date", auth)

    def test_credentials_env_dict(self):
        creds = aws_cp_gui.AWSCredentialsConfig(
            access_key_id="AKIA123",
            secret_access_key="SECRET456",
            region="ap-southeast-2",
            output_format="table",
        )
        env = creds.to_env_dict()
        self.assertEqual(env["AWS_ACCESS_KEY_ID"], "AKIA123")
        self.assertEqual(env["AWS_SECRET_ACCESS_KEY"], "SECRET456")
        self.assertEqual(env["AWS_DEFAULT_REGION"], "ap-southeast-2")
        self.assertEqual(env["AWS_DEFAULT_OUTPUT"], "table")

    def test_gui_instantiation(self):
        root = tk.Tk()
        root.withdraw()
        try:
            app = aws_cp_gui.AwsCpTransferApp(root)
            app.var_access_key.set("AKIA_TEST")
            app.var_secret_key.set("SECRET_TEST")
            app.var_region.set("eu-west-1")
            app.var_output_format.set("yaml")
            app.var_source_path.set(r"C:\temp\file.txt")
            app.var_target_path.set("s3://my-bucket/dest/")
            app.update_command_preview()
            self.assertIn("--region eu-west-1", app.var_cmd_preview.get())
            self.assertIn("--output yaml", app.var_cmd_preview.get())
        finally:
            root.destroy()

    def test_powershell_script_variables_and_syntax(self):
        ps1_path = Path(__file__).resolve().parent / "aws_cp_gui.ps1"
        content = ps1_path.read_text(encoding="utf-8")
        self.assertIn('${url}?${canonicalQuery}', content)
        self.assertNotIn('"$url?$canonicalQuery"', content)


if __name__ == "__main__":
    unittest.main()
