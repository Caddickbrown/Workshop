#!/bin/bash

# Display Settings

# Set the default resolution to 720p and refresh rate to 30Hz
echo "Setting default resolution to 720p and refresh rate to 30Hz..."
sudo sed -i '/^hdmi_group=/d' /boot/config.txt
sudo sed -i '/^hdmi_mode=/d' /boot/config.txt
sudo sed -i '/^hdmi_force_hotplug=/d' /boot/config.txt

echo "hdmi_group=1" | sudo tee -a /boot/config.txt
echo "hdmi_mode=4" | sudo tee -a /boot/config.txt
echo "hdmi_force_hotplug=1" | sudo tee -a /boot/config.txt

# Explanation:
# - hdmi_group=1: CEA (Consumer Electronics Association) group, used for TVs.
# - hdmi_mode=4: 720p resolution at 60Hz (by default). We need to override this to 30Hz.
# - hdmi_force_hotplug=1: Forces the HDMI mode even if no HDMI monitor is detected.

# Set the refresh rate to 30Hz
echo "Setting refresh rate to 30Hz..."
sudo sed -i '/^hdmi_mode/!b;n;c\hdmi_mode=2' /boot/config.txt
hdmi_mode=2 corresponds to 720p at 30Hz.

# Clean up
echo "Cleaning up..."
sudo apt-get autoremove -y
sudo apt-get clean

# Reboot the system
echo "Rebooting the system..."
sudo reboot
