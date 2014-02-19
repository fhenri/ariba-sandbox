# -*- mode: ruby -*-
# vi: set ft=ruby :

# Vagrantfile API/syntax version. Don't touch unless you know what you're doing!
VAGRANTFILE_API_VERSION = "2"

Vagrant.configure(VAGRANTFILE_API_VERSION) do |config|

  # Every Vagrant virtual environment requires a box to build off of.
  config.vm.box = "aribabox"
  config.ssh.username = "ariba"

  config.vm.provider :virtualbox do |vb|
    # Don't boot with headless mode
    #vb.gui = true

    # update memory - keep to 3Gb but if all 3 instances are started, bring up more memory!
    vb.customize ["modifyvm", :id, "--memory", "3024"]
  end

  config.vm.define "sourcing" do |sourcing|
    sourcing.vm.network :private_network, ip: "192.168.60.10"
  end

  config.vm.define "buyer" do |buyer|
    buyer.vm.network :private_network, ip: "192.168.60.12"
  end

  config.vm.define "asm" do |asm|
    asm.vm.network :private_network, ip: "192.168.60.16"
  end

end
