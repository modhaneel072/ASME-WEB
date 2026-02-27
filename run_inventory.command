#!/bin/bash
cd "$(dirname "$0")"

# Load conda so "conda activate" works in a double-click script
source /opt/anaconda3/etc/profile.d/conda.sh

conda activate asme_inventory

python app.py &
sleep 1
open "http://127.0.0.1:5000"