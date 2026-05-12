from pydantic_settings import BaseSettings

class Settings(BaseSettings):
    # Database - mặc định SQLite để chạy local, khi deploy sẽ bị ghi đè bởi biến môi trường
    DATABASE_URL: str = "sqlite:///./data/app.db"
    SECRET_KEY: str = "default-secret-change-in-production"
    ALGORITHM: str = "HS256"
    ACCESS_TOKEN_EXPIRE_MINUTES: int = 60
    UPLOAD_DIR: str = "uploads"

    class Config:
        env_file = ".env"

settings = Settings()