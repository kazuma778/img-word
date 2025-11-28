"""
Logging setup untuk aplikasi Image Processor
Log akan disimpan di log.txt dan dibersihkan setiap restart
"""
import logging
import os


def setup_file_logging():
    """Setup logging ke file dan console"""
    # Path ke log file
    log_file = os.path.join(os.path.dirname(__file__), 'log.txt')
    
    # Remove existing log handlers to avoid duplicates
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    
    # Configure logging with both console and file handlers
    handlers = [logging.StreamHandler()]
    
    try:
        # File handler - mode 'w' clears file on each startup
        handlers.append(logging.FileHandler(log_file, mode='w', encoding='utf-8'))
    except OSError:
        # Fallback for read-only file systems (like Vercel)
        pass

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(name)s - %(message)s',
        handlers=handlers
    )
    
    # Return logger for app to use
    return logging.getLogger('img_word')
