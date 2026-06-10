#!/usr/bin/env python3
"""
Zerodha Training Hub — Background Database Watcher
Monitors Firebase database for changes and automatically regenerates the styled Excel report.
"""

import time
import urllib.request
import json
import subprocess
import os
import hashlib
import sys

FIREBASE_URL = "https://zerodha-training-hub-default-rtdb.firebaseio.com/zh_logs.json"
SCRIPT_PATH  = "/Users/girisha/Desktop/Zerodha-training-hub/generate_zerodha_report.py"
LOG_PATH     = "/Users/girisha/Desktop/Zerodha-training-hub/watcher.log"

def log_msg(msg):
    t = time.strftime('%Y-%m-%d %H:%M:%S')
    formatted = f"[{t}] {msg}"
    print(formatted)
    try:
        with open(LOG_PATH, "a") as f:
            f.write(formatted + "\n")
    except:
        pass

def get_data_hash():
    try:
        req = urllib.request.Request(
            FIREBASE_URL, 
            headers={'User-Agent': 'Mozilla/5.0'}
        )
        with urllib.request.urlopen(req, timeout=10) as r:
            data = r.read()
            return hashlib.md5(data).hexdigest()
    except Exception as e:
        log_msg(f"Error checking Firebase: {e}")
        return None

def main():
    log_msg("Watcher service started. Monitoring Firebase...")
    last_hash = get_data_hash()
    
    # Run initially to make sure Desktop file is up to date
    if last_hash:
        log_msg("Initial regeneration check...")
        try:
            subprocess.run(["python3", SCRIPT_PATH], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            log_msg("Excel sheet updated.")
        except Exception as e:
            log_msg(f"Initial regeneration failed: {e}")

    while True:
        try:
            time.sleep(10)
            current_hash = get_data_hash()
            if current_hash and current_hash != last_hash:
                log_msg("Database change detected! Regenerating styled Excel report...")
                try:
                    res = subprocess.run(
                        ["python3", SCRIPT_PATH], 
                        check=True, 
                        stdout=subprocess.PIPE, 
                        stderr=subprocess.PIPE,
                        text=True
                    )
                    log_msg("Styled report updated successfully on Desktop.")
                    # Log snippet of output
                    for line in res.stdout.splitlines():
                        if "Saved" in line or "Sheets" in line:
                            log_msg(f"Generator output: {line.strip()}")
                except subprocess.CalledProcessError as e:
                    log_msg(f"Failed to run generator: {e.stderr.strip() if e.stderr else e}")
                except Exception as e:
                    log_msg(f"Failed to regenerate report: {e}")
                
                last_hash = current_hash
        except KeyboardInterrupt:
            log_msg("Watcher service stopped by keyboard interrupt.")
            break
        except Exception as e:
            log_msg(f"Watcher loop error: {e}")

if __name__ == "__main__":
    main()
