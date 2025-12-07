#!/usr/bin/env bash
apt-get update
apt-get install -y wget gnupg unzip

# Install Google Chrome
wget https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb
apt-get install -y ./google-chrome-stable_current_amd64.deb

echo "Chrome installed successfully"
