#!/bin/bash
cd /home/main/Documents/BAE_Code
source /home/main/BAEvenv/bin/activate
export DISPLAY=:0
export XAUTHORITY=/home/main/.Xauthority
python BAE_SW_Code.py
