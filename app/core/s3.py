import boto3
from app.core.config import settings
from app.core.logging import get_logger

logger = get_logger(__name__)

def get_s3_client():
    """Create and return a boto3 S3 client configured with MinIO settings."""
    return boto3.client(
        "s3",
        endpoint_url=settings.S3_ENDPOINT,
        aws_access_key_id=settings.S3_ACCESS_KEY_ID,
        aws_secret_access_key=settings.S3_ACCESS_SECRET,
        region_name=settings.S3_REGION,
        verify=settings.S3_VERIFY,
    )

def upload_to_s3(content: bytes, key: str) -> None:
    """Upload raw bytes to S3/MinIO bucket."""
    s3 = get_s3_client()
    try:
        s3.put_object(
            Bucket=settings.S3_BUCKET,
            Key=key,
            Body=content
        )
        logger.info(f"Uploaded file to S3: Bucket={settings.S3_BUCKET}, Key={key}")
    except Exception as e:
        logger.error(f"Failed to upload file to S3: {str(e)}")
        raise
