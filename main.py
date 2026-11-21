import os
import time
import json
import logging
import threading
from datetime import datetime
import psutil
import win32gui
import win32process
import win32com .client
class ActivityMonitor :
    def __init__ (self ,config_path ='config.json'):
        self .config =self ._load_config (config_path )
        self ._setup_logging ()
        self .current_path =None
        self .tracked_pids =self ._get_current_pids ()
        self .stop_event =threading .Event ()
    def _load_config (self ,config_path ):
        with open (config_path ,'r',encoding ='utf-8')as f :
            return json .load (f )
    def _setup_logging (self ):
        self .activity_logger =self ._create_logger (
        'activity',
        self .config ['log_file'],
        logging .INFO
        )
        self .error_logger =self ._create_logger (
        'error',
        self .config ['error_log'],
        logging .ERROR
        )
    def _create_logger (self ,name ,filename ,level ):
        logger =logging .getLogger (name )
        logger .setLevel (level )
        handler =logging .FileHandler (filename ,encoding ='utf-8')
        formatter =logging .Formatter ('%(asctime)s - %(сообщение)s',"%Y-%m-%d %H:%M:%S")
        handler .setFormatter (formatter )
        logger .addHandler (handler )
        return logger
    def _get_current_pids (self ):
        return {p .pid for p in psutil .process_iter (['pid'])}
    def _monitor_explorer_path (self ):
        shell =win32com .client .Dispatch ("Shell.Application")
        while not self .stop_event .is_set ():
            try :
                active_window_handle =win32gui .GetForegroundWindow ()
                _ ,pid =win32process .GetWindowThreadProcessId (active_window_handle )
                process_name =psutil .Process (pid ).name ()
                if "explorer.exe"in process_name .lower ():
                    for window in shell .Windows ():
                        if window .HWND ==active_window_handle :
                            path =window .Document .Folder .Self .Path
                            if path !=self .current_path :
                                self .current_path =path
                                self .activity_logger .info (f"Entered directory: {self .current_path }")
                            break
            except (psutil .NoSuchProcess ,psutil .AccessDenied ):
                self .current_path =None
            except Exception :
                self .current_path =None
            time .sleep (self .config ['directory_check_interval_sec'])
    def _monitor_new_processes (self ):
        while not self .stop_event .is_set ():
            try :
                current_pids =self ._get_current_pids ()
                new_pids =current_pids -self .tracked_pids
                for pid in new_pids :
                    try :
                        proc =psutil .Process (pid )
                        self .activity_logger .info (f"New process: {proc .name ()} (PID: {pid })")
                    except (psutil .NoSuchProcess ,psutil .AccessDenied ):
                        continue
                self .tracked_pids =current_pids
            except Exception as e :
                self .error_logger .error (f"Error in process monitor: {e }",exc_info =True )
            time .sleep (self .config ['process_check_interval_sec'])
    def start (self ):
        self .activity_logger .info ("Монитор активности запустился.")
        dir_thread =threading .Thread (target =self ._monitor_explorer_path ,daemon =True )
        proc_thread =threading .Thread (target =self ._monitor_new_processes ,daemon =True )
        dir_thread .start ()
        proc_thread .start ()
        try :
            while True :
                time .sleep (1 )
        except KeyboardInterrupt :
            self .stop_event .set ()
            self .activity_logger .info ("Монитор активности остановлен.")
if __name__ =="__main__":
    monitor =ActivityMonitor ()
    monitor .start ()