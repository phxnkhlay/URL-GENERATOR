#!/usr/bin/env python3
"""
generate_all_repos_raw_urls.py

Ambil SEMUA FILE dari seluruh repository GitHub milik satu user/org,
kemudian:

1. Membaca NAMA FILE / PATH FILE
2. Mencari tanggal + bulan + tahun di dalam nama file
3. HANYA mengambil file yang periodenya tahun 2027
4. Mengurutkan:
      JANUARI
      FEBRUARI
      MARET
      APRIL
      MEI
      JUNI
      JULI
      AGUSTUS
      SEPTEMBER
      OKTOBER
      NOVEMBER
      DESEMBER
5. Di dalam bulan diurutkan tanggal 01 -> 31
6. Prefix file/repository TIDAK menjadi patokan
7. Menyimpan hasil ke CSV dan XLSX

Contoh nama file yang dikenali:

    AL01JANUARI2027
    AP02JANUARI2027
    ABC15FEBRUARI2027
    XX31MARET2027
    FILE01DESEMBER2027

Prefix AL / AP / ABC / XX / dll tidak berpengaruh.

Cara menjalankan:

    python generate_all_repos_raw_urls.py phxnkhlay

Output:

    phxnkhlay-2027-raw_urls.csv
    phxnkhlay-2027-raw_urls.xlsx

Private repository:

Windows CMD:
    set GITHUB_TOKEN=ghp_xxxxxxxxx

PowerShell:
    $env:GITHUB_TOKEN="ghp_xxxxxxxxx"

Linux/Mac:
    export GITHUB_TOKEN="ghp_xxxxxxxxx"

Lalu:

    python generate_all_repos_raw_urls.py phxnkhlay
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import time

from typing import Any, Dict, List, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests


# ============================================================
# CONFIG
# ============================================================

GITHUB_API = "https://api.github.com"
PER_PAGE = 100

TARGET_YEAR = 2027

MONTHS = {
    "JANUARI": 1,
    "FEBRUARI": 2,
    "MARET": 3,
    "APRIL": 4,
    "MEI": 5,
    "JUNI": 6,
    "JULI": 7,
    "AGUSTUS": 8,
    "SEPTEMBER": 9,
    "OKTOBER": 10,
    "NOVEMBER": 11,
    "DESEMBER": 12,
}


# ============================================================
# GITHUB REQUEST
# ============================================================

def github_get(
    url: str,
    token: Optional[str] = None,
    params: Dict[str, Any] | None = None,
    timeout: int = 30,
):
    headers = {
        "Accept": "application/vnd.github.v3+json"
    }

    if token:
        headers["Authorization"] = f"token {token}"

    r = requests.get(
        url,
        headers=headers,
        params=params,
        timeout=timeout
    )

    # Handle GitHub rate limit
    if (
        r.status_code == 403
        and r.headers.get("X-RateLimit-Remaining") == "0"
    ):
        reset = int(
            r.headers.get(
                "X-RateLimit-Reset",
                time.time() + 60
            )
        )

        wait = max(
            5,
            reset - int(time.time()) + 3
        )

        print(
            f"[WARN] Rate limited. "
            f"Tidur {wait} detik...",
            file=sys.stderr
        )

        time.sleep(wait)

        r = requests.get(
            url,
            headers=headers,
            params=params,
            timeout=timeout
        )

    return r


# ============================================================
# LIST ALL REPOSITORIES
# ============================================================

def list_repos_for_user(
    owner: str,
    token: Optional[str] = None
) -> List[Dict[str, Any]]:

    repos: List[Dict[str, Any]] = []

    page = 1

    while True:

        url = f"{GITHUB_API}/users/{owner}/repos"

        params = {
            "per_page": PER_PAGE,
            "page": page,
            "type": "all",
            "sort": "full_name",
            "direction": "asc",
        }

        r = github_get(
            url,
            token=token,
            params=params
        )

        if r.status_code != 200:
            raise RuntimeError(
                f"Gagal list repos {owner}: "
                f"{r.status_code} {r.text}"
            )

        items = r.json()

        if not items:
            break

        repos.extend(items)

        if len(items) < PER_PAGE:
            break

        page += 1

    return repos


# ============================================================
# GET REPOSITORY TREE
# ============================================================

def get_tree_recursive(
    owner: str,
    repo: str,
    branch: str,
    token: Optional[str] = None
) -> Dict[str, Any]:

    url = (
        f"{GITHUB_API}/repos/"
        f"{owner}/{repo}/git/trees/{branch}"
    )

    params = {
        "recursive": 1
    }

    r = github_get(
        url,
        token=token,
        params=params
    )

    # Fallback branch reference
    if r.status_code == 404:

        url2 = (
            f"{GITHUB_API}/repos/"
            f"{owner}/{repo}/git/trees/"
            f"refs/heads/{branch}"
        )

        r = github_get(
            url2,
            token=token,
            params=params
        )

    if r.status_code != 200:
        raise RuntimeError(
            f"Gagal ambil tree "
            f"{owner}/{repo}@{branch}: "
            f"{r.status_code} {r.text}"
        )

    return r.json()


# ============================================================
# EXTRACT TANGGAL DARI NAMA FILE
# ============================================================

def extract_date_from_filename(
    path: str
) -> Optional[tuple[int, int, int]]:

    """
    Mencari pola:

        DDMMMMYYYY

    Contoh:

        AL01JANUARI2027
        AP15FEBRUARI2027
        TEST31DESEMBER2027

    Prefix apa pun diperbolehkan.

    Hanya tahun 2027 yang diterima.
    """

    if not path:
        return None

    # Gunakan nama file/path sebagai sumber data.
    #
    # Contoh:
    # folder/AL01JANUARI2027.m3u
    #
    # tetap akan terbaca.

    name = os.path.basename(path).upper()

    pattern = re.compile(
        r"(?<!\d)"
        r"(\d{1,2})"
        r"(JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|"
        r"AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)"
        r"(2027)"
        r"(?!\d)"
    )

    match = pattern.search(name)

    if not match:
        return None

    day = int(match.group(1))

    month_name = match.group(2)

    year = int(match.group(3))

    month = MONTHS[month_name]

    # Validasi tanggal
    if day < 1 or day > 31:
        return None

    return (
        year,
        month,
        day
    )


# ============================================================
# PROCESS SATU REPOSITORY
# ============================================================

def build_rows_for_repo(
    owner: str,
    repo_item: Dict[str, Any],
    token: Optional[str] = None
) -> List[Dict[str, Any]]:

    repo_name = repo_item.get("name")

    branch = (
        repo_item.get("default_branch")
        or "main"
    )

    try:

        tree_json = get_tree_recursive(
            owner,
            repo_name,
            branch,
            token=token
        )

    except Exception as e:

        print(
            f"[ERROR] {repo_name}: {e}",
            file=sys.stderr
        )

        return []

    rows: List[Dict[str, Any]] = []

    for node in tree_json.get("tree", []):

        # Hanya file
        if node.get("type") != "blob":
            continue

        path = node.get("path")

        if not path:
            continue

        # ====================================================
        # PENTING:
        # TANGGAL DIAMBIL DARI NAMA FILE
        # BUKAN NAMA REPOSITORY
        # ====================================================

        date_info = extract_date_from_filename(
            path
        )

        # Tidak ada tanggal 2027 -> skip
        if date_info is None:
            continue

        year, month, day = date_info

        raw_url = (
            f"https://raw.githubusercontent.com/"
            f"{owner}/"
            f"{repo_name}/"
            f"{branch}/"
            f"{path}"
        )

        rows.append(
            {
                "repo": repo_name,
                "file": os.path.basename(path),
                "path": path,
                "size": node.get("size"),
                "branch": branch,
                "year": year,
                "month": month,
                "day": day,
                "raw_url": raw_url,
            }
        )

    return rows


# ============================================================
# SORTING
# ============================================================

def sort_rows(
    rows: List[Dict[str, Any]]
) -> List[Dict[str, Any]]:

    """
    Urutan:

        Tahun
        Bulan
        Tanggal
        Nama file
        Repository
        Path
    """

    return sorted(
        rows,
        key=lambda row: (
            row.get("year", 9999),
            row.get("month", 99),
            row.get("day", 99),
            (row.get("file") or "").upper(),
            (row.get("repo") or "").upper(),
            (row.get("path") or "").upper(),
        )
    )


# ============================================================
# MAIN
# ============================================================

def main():

    parser = argparse.ArgumentParser(
        description=(
            "Ambil raw URL file dari seluruh "
            "repository GitHub dan urutkan "
            "berdasarkan tanggal Januari-Desember 2027 "
            "yang ditemukan di NAMA FILE."
        )
    )

    parser.add_argument(
        "owner",
        help="Username / Organization GitHub"
    )

    parser.add_argument(
        "--token",
        default=os.getenv("GITHUB_TOKEN"),
        help=(
            "GitHub PAT. "
            "Bisa menggunakan GITHUB_TOKEN."
        )
    )

    parser.add_argument(
        "--out",
        default=None,
        help=(
            "Nama output CSV. "
            "Default: <owner>-2027-raw_urls.csv"
        )
    )

    parser.add_argument(
        "--workers",
        type=int,
        default=6,
        help="Jumlah proses paralel. Default 6."
    )

    parser.add_argument(
        "--only-public",
        action="store_true",
        help="Hanya proses repository public."
    )

    args = parser.parse_args()

    owner = args.owner

    token = args.token

    out_name = (
        args.out
        or f"{owner}-2027-raw_urls.csv"
    )


    # ========================================================
    # LIST REPOSITORIES
    # ========================================================

    print()
    print("=" * 60)
    print(
        f"Mencari seluruh repository milik: {owner}"
    )
    print("=" * 60)

    repos = list_repos_for_user(
        owner,
        token=token
    )

    if args.only_public:

        repos = [
            r for r in repos
            if not r.get("private", False)
        ]

    print(
        f"Ditemukan {len(repos)} repository."
    )

    print(
        f"Workers: {args.workers}"
    )

    print()


    # ========================================================
    # PROCESS SEMUA REPOSITORY
    # ========================================================

    all_rows: List[
        Dict[str, Any]
    ] = []

    with ThreadPoolExecutor(
        max_workers=args.workers
    ) as executor:

        future_map = {
            executor.submit(
                build_rows_for_repo,
                owner,
                repo,
                token
            ): repo
            for repo in repos
        }

        for future in as_completed(
            future_map
        ):

            repo = future_map[future]

            repo_name = repo.get(
                "name",
                "UNKNOWN"
            )

            try:

                rows = future.result()

                all_rows.extend(
                    rows
                )

                print(
                    f"[OK] "
                    f"{repo_name:<35} "
                    f"{len(rows)} file 2027"
                )

            except Exception as e:

                print(
                    f"[ERROR] "
                    f"{repo_name}: {e}",
                    file=sys.stderr
                )


    # ========================================================
    # CHECK HASIL
    # ========================================================

    if not all_rows:

        print()
        print(
            "TIDAK ADA FILE 2027 "
            "YANG DITEMUKAN."
        )

        print(
            "Pastikan nama file mengandung pola "
            "seperti:"
        )

        print(
            "AL01JANUARI2027"
        )

        sys.exit(0)


    # ========================================================
    # SORT
    # ========================================================

    print()
    print(
        "Mengurutkan file berdasarkan "
        "JANUARI -> DESEMBER 2027..."
    )

    all_rows = sort_rows(
        all_rows
    )


    # ========================================================
    # DATAFRAME
    # ========================================================

    df = pd.DataFrame(
        all_rows
    )

    columns = [
        "repo",
        "file",
        "path",
        "size",
        "branch",
        "year",
        "month",
        "day",
        "raw_url",
    ]

    df = df[
        columns
    ]


    # ========================================================
    # CSV PATH
    # ========================================================

    if out_name.lower().endswith(
        ".csv"
    ):

        csv_path = out_name

    else:

        csv_path = (
            os.path.splitext(
                out_name
            )[0]
            + ".csv"
        )

    os.makedirs(
        os.path.dirname(csv_path)
        or ".",
        exist_ok=True
    )


    # ========================================================
    # SAVE CSV
    # ========================================================

    df.to_csv(
        csv_path,
        index=False,
        encoding="utf-8-sig"
    )

    print()
    print(
        f"CSV berhasil dibuat:"
    )

    print(
        csv_path
    )


    # ========================================================
    # SAVE XLSX
    # ========================================================

    xlsx_path = (
        os.path.splitext(
            csv_path
        )[0]
        + ".xlsx"
    )

    try:

        df.to_excel(
            xlsx_path,
            index=False
        )

        print(
            "Excel berhasil dibuat:"
        )

        print(
            xlsx_path
        )

    except Exception as e:

        print(
            f"[WARN] Gagal membuat XLSX: {e}",
            file=sys.stderr
        )


    # ========================================================
    # SUMMARY
    # ========================================================

    print()
    print("=" * 60)
    print("RINGKASAN")
    print("=" * 60)

    print(
        f"Total file 2027 : {len(df)}"
    )

    print(
        f"Januari         : "
        f"{len(df[df['month'] == 1])}"
    )

    print(
        f"Februari        : "
        f"{len(df[df['month'] == 2])}"
    )

    print(
        f"Maret           : "
        f"{len(df[df['month'] == 3])}"
    )

    print(
        f"April           : "
        f"{len(df[df['month'] == 4])}"
    )

    print(
        f"Mei             : "
        f"{len(df[df['month'] == 5])}"
    )

    print(
        f"Juni            : "
        f"{len(df[df['month'] == 6])}"
    )

    print(
        f"Juli            : "
        f"{len(df[df['month'] == 7])}"
    )

    print(
        f"Agustus         : "
        f"{len(df[df['month'] == 8])}"
    )

    print(
        f"September       : "
        f"{len(df[df['month'] == 9])}"
    )

    print(
        f"Oktober         : "
        f"{len(df[df['month'] == 10])}"
    )

    print(
        f"November        : "
        f"{len(df[df['month'] == 11])}"
    )

    print(
        f"Desember        : "
        f"{len(df[df['month'] == 12])}"
    )

    print("=" * 60)


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":
    main()
