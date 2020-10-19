#!/usr/bin/env python3 
# -*- coding: utf-8 -*-
import os
import time
while True:
    os.system('taskkill /im chrome.exe /f')
    time.sleep(3600)
