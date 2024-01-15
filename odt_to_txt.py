#!/usr/bin/env python3

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import textwrap
from datetime import datetime
from pathlib import Path
from typing import NamedTuple

WRAP_WIDTH = 112


class AppOptions(NamedTuple):
    path_list: list[str]
    do_recurse: bool
    do_overwrite: bool
    dt_tag: bool


warnings = []


def get_args(argv):
    ap = argparse.ArgumentParser(
        description="Run LibreOffice to convert document files to text files. "
        "Handles .odt, .doc, and .docx file formats."
    )

    ap.add_argument(
        "paths",
        nargs="*",
        action="store",
        help="One or more files and/or directories to process. Files must "
        "be type (extension) '.odt', '.doc', or '.docx'. For directories, "
        "all files with one of those extensions will be processed.",
    )

    ap.add_argument(
        "-r",
        "--recurse",
        dest="do_recurse",
        action="store_true",
        help="Do recursive search for document files in sub-directories.",
    )

    ap.add_argument(
        "-o",
        "--overwrite",
        dest="do_overwrite",
        action="store_true",
        help="Overwrite existing output files. By default, existing files are "
        "not replaced.",
    )

    ap.add_argument(
        "-d",
        "--datetime-tag",
        dest="dt_tag",
        action="store_true",
        help="Add a [date_time] tag, based on the source document last "
        "modified timestamp, to the output file names.",
    )

    return ap.parse_args(argv[1:])


def get_options(argv) -> AppOptions:
    args = get_args(argv)
    return AppOptions(args.paths, args.do_recurse, args.do_overwrite, args.dt_tag)


def get_date_time_tag(file: Path) -> str:
    assert file.exists()  # noqa: S101
    dt = datetime.fromtimestamp(file.stat().st_mtime)
    return dt.strftime("%Y%m%d_%H%M")


def make_wrapped_version(source_path: Path, opts: AppOptions):
    assert source_path.exists()  # noqa: S101

    target_path = source_path.parent / f"{source_path.stem}-wrap.txt"

    # if target_path.exists() and not opts.do_overwrite:
    #     warnings.append(f"ALREADY EXISTS: '{target_path}'")
    #     return

    print(f"Wrap: '{target_path}'")
    with target_path.open("w") as w, source_path.open() as r:
        for line_in in r.readlines():
            wrapped = textwrap.wrap(line_in, width=WRAP_WIDTH)
            if wrapped:
                for line_out in wrapped:
                    w.write(f"{line_out}\n")
            else:
                w.write("\n")


def run_convert_to_txt(in_path: Path, opts: AppOptions):
    assert isinstance(in_path, Path)  # noqa: S101

    print(f"ODT: '{in_path}'")

    if opts.dt_tag:
        new_name = str(
            in_path.with_suffix(f"{in_path.suffix}-{get_date_time_tag(in_path)}-as.txt")
        )
    else:
        new_name = str(in_path.with_suffix(f"{in_path.suffix}-as.txt"))

    if Path(new_name).exists() and not opts.do_overwrite:
        warnings.append(f"SKIP EXISTING '{new_name}'")
        warnings.append(
            "  Existing files are not overwritten unless the --overwrite "
            "(-o) option is used."
        )
        return

    exe = shutil.which("libreoffice")

    cmd = [
        exe,
        "--convert-to",
        "txt",
        "--outdir",
        str(in_path.parent),
        str(in_path),
    ]

    subprocess.check_call(cmd, stderr=subprocess.DEVNULL)  # noqa: S603

    out_path = in_path.with_suffix(".txt")
    assert out_path.exists()  # noqa: S101

    print(f"  as: '{new_name}'")

    shutil.move(str(out_path), new_name)

    make_wrapped_version(Path(new_name), opts)


def get_files(dir_path: Path, file_ext: str, opts: AppOptions) -> list[Path]:
    if opts.do_recurse:
        files = sorted(dir_path.glob(f"**/*{file_ext}"))
    else:
        files = sorted(
            [
                f
                for f in dir_path.iterdir()
                if f.is_file() and f.suffix.lower() == file_ext
            ]
        )
    return files


def process_paths(opts: AppOptions):  # noqa: PLR0912
    for name in opts.path_list:
        p = Path(name)

        if not p.exists():
            warnings.append(f"Path not found: '{p}'")
            continue

        if p.is_file():
            if p.suffix.lower() not in [".odt", ".doc", ".docx"]:
                warnings.append("Not a supported file type: '{p}'")
                continue
            run_convert_to_txt(p, opts)

        elif p.is_dir():
            files = get_files(p, ".odt", opts)
            for f in files:
                run_convert_to_txt(f, opts)

            files = get_files(p, ".docx", opts)
            for f in files:
                run_convert_to_txt(f, opts)

            files = get_files(p, ".doc", opts)
            for f in files:
                run_convert_to_txt(f, opts)

            files = get_files(p, ".bak", opts)
            for f in files:
                if ".odt" in f.name or ".doc" in f.name:
                    #  '.doc' also matches '.docx'
                    run_convert_to_txt(f, opts)
                else:
                    print(f"SKIP: {f}")
        else:
            warnings.append(f"Cannot process path '{p}'.")


def main(argv):
    opts = get_options(argv)
    process_paths(opts)

    if warnings:
        print("\nWARNINGS:")
        for warning in warnings:
            print(warning)

    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
