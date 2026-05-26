import io
import zipfile
import urllib.request

import pandas as pd

from src.config import DATA_DIR, RAW_DATA_PATH


UCI_ZIP_URL = "https://archive.ics.uci.edu/static/public/222/bank+marketing.zip"


def download_dataset() -> None:
    DATA_DIR.mkdir(exist_ok=True)
    if RAW_DATA_PATH.exists():
        return

    zip_path = DATA_DIR / "bank_marketing.zip"
    urllib.request.urlretrieve(UCI_ZIP_URL, zip_path)

    with zipfile.ZipFile(zip_path) as outer_zip:
        names = outer_zip.namelist()
        if "bank-additional/bank-additional-full.csv" in names:
            with outer_zip.open("bank-additional/bank-additional-full.csv") as src:
                RAW_DATA_PATH.write_bytes(src.read())
            return

        nested_bytes = outer_zip.read("bank-additional.zip")
        with zipfile.ZipFile(io.BytesIO(nested_bytes)) as nested_zip:
            with nested_zip.open("bank-additional/bank-additional-full.csv") as src:
                RAW_DATA_PATH.write_bytes(src.read())


def load_dataset() -> pd.DataFrame:
    download_dataset()
    return pd.read_csv(RAW_DATA_PATH, sep=";")


def split_features_target(df: pd.DataFrame, include_duration: bool = False):
    y = (df["y"] == "yes").astype(int)
    X = df.drop(columns=["y"])
    if not include_duration and "duration" in X.columns:
        X = X.drop(columns=["duration"])
    return X, y
