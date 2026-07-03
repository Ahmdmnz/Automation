from pydantic_settings import BaseSettings, SettingsConfigDict
from functools import lru_cache

class Settings(BaseSettings):
    GROQ_API_KEYS: str
    EXCEL_FILE_PATH: str
    IMAGES_FOLDER: str
    AI_MODEL: str = "llama-3.3-70b-versatile"

    # Batch control
    START_ROW: int = 1
    MAX_ROWS_PER_RUN: int | None = None

    # Sheet
    SHEET_NAME: str = "Sheet1"

    # Input columns
    ITEM_NUMBER_COLUMN: str = "D"
    BARCODE_COLUMN: str = "F"

    # Scraped data columns
    COL_NAME_EN:  str = "H"
    COL_DESC_EN:  str = "J"
    COL_SPECS_EN: str = "L"

    # Output columns
    COL_META_TITLE_EN: str = "N"
    COL_META_DESC_EN:  str = "P"
    COL_TAGS_EN:       str = "R"
    COL_BRAND_EN:      str = "T"

    model_config = SettingsConfigDict(env_file=".env")

@lru_cache
def get_settings():
    return Settings()