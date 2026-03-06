from dataclasses import dataclass
import os
from dotenv import load_dotenv

load_dotenv()


@dataclass(frozen=True)
class Settings:
    bb_client_id: str
    bb_client_secret: str
    bb_developer_key: str
    bb_token_url: str
    bb_api_base_url: str
    bb_extrato_path: str
    bb_scope: str



def load_settings() -> Settings:
    return Settings(
        bb_client_id=os.getenv("BB_CLIENT_ID", ""),
        bb_client_secret=os.getenv("BB_CLIENT_SECRET", ""),
        bb_developer_key=os.getenv("BB_DEVELOPER_KEY", ""),
        bb_token_url=os.getenv("BB_TOKEN_URL", "https://oauth.bb.com.br/oauth/token"),
        bb_api_base_url=os.getenv("BB_API_BASE_URL", "https://api.bb.com.br"),
        bb_extrato_path=os.getenv("BB_EXTRATO_PATH", "/extratos/v1/conta-corrente/extrato"),
        bb_scope=os.getenv("BB_SCOPE", "extrato-info"),
    )
