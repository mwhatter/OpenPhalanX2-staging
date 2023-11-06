#!/bin/bash

# Enable IP forwarding permanently
echo "Enabling IP forwarding..."
echo 'net.ipv4.ip_forward=1' | sudo tee -a /etc/sysctl.conf > /dev/null
sudo sysctl -p /etc/sysctl.conf

# Install iptables-services if not already present
echo "Installing iptables-services..."
sudo yum install -y iptables-services

# Enable iptables service to start on boot
sudo systemctl enable iptables

# Set up IP masquerading for interfaces eth1 through eth8
for i in {1..8}; do
  echo "Setting up IP masquerading for eth$i..."
  sudo iptables -t nat -A POSTROUTING -o eth$i -j MASQUERADE
done

# Save the iptables rule
echo "Saving iptables rules..."
sudo service iptables save

echo "IP forwarding and masquerading have been configured for interfaces eth1 through eth8."

# END OF SCRIPT

file.io/9pkcy9iicK2D
