#!/usr/bin/env python
import sys
import time
import socket

def is_port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        try:
            s.bind(("127.0.0.1", port))
            return False
        except:
            return True

def main():
    from app import app
    from werkzeug.serving import run_simple
    import logging
    
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger(__name__)
    
    logger.info("=" * 60)
    logger.info("Starting Markdown to Word Converter")
    logger.info("=" * 60)
    logger.info("Flask server starting...")
    logger.info("Open: http://localhost:5000")
    logger.info("=" * 60)
    
    try:
        run_simple(
            "0.0.0.0",
            5000,
            app,
            use_debugger=False,
            use_reloader=False,
            threaded=True,
            static_files={"/static": "static"}
        )
    except KeyboardInterrupt:
        logger.info("Server stopped")
        sys.exit(0)
    except Exception as e:
        logger.error(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
