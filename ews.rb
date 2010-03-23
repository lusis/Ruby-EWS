#!/usr/bin/ruby

# EWS - Exchange Web Services interface
# Copyright (C) 2009  John E. Vincent <rubystuff@lusis.org>.

# This program is copyrighted free software by John E. Vincent.  You can
# redistribute it and/or modify it under the same terms of Ruby's license;
# either the dual license version in 2003, or any later version.
require 'optparse'
require 'ewslookup.rb'
require 'pp'

user='myusername' # Username needed for access
pass='mypassword' # Password needed for access
endpoint = "https://my.exchange.server/EWS/Exchange.asmx" # URL to Exchange.asmx

OPTIONS = {}
OptionParser.new do |opts|
    script_name = File.basename($0)
    opts.banner = "Usage: #{script_name} [options] <entry>"
    opts.define_head    "Ruby script to look up address book (GAL, DL and Private) entries from an Exchange server.
			This uses the EWS interface of Exchange Server 2007."

    opts.on("-v", "--verbose", "Run Verbosely") do |v|
        OPTIONS[:verbose] = v
    end
    opts.on("-d", "--debug", "Enable debug mode") do |d|
        OPTIONS[:debug] = d
    end
    opts.on("-g", "--gal query", "Search entires in the Exchange GAL.") do |g|
		OPTIONS[:gal] = g
    end
    opts.on("-l", "--distlist query", String, "Get details members of a given distribution list.") do |l|
		OPTIONS[:distlist] = l
    end
    opts.on_tail("-h", "--help", "Show this message") do
        puts opts
        exit
    end
end.parse!

ews = EWSLookup.new(user,pass,endpoint)
res = ews.gal_lookup(OPTIONS[:gal]) if OPTIONS[:gal]
res = ews.dl_lookup(OPTIONS[:distlist]) if OPTIONS[:distlist]
pp res

# Some random fun facts
# Count of members in DL
#pp resp2[0].dLExpansion.xmlattr_TotalItemsInView

# Name of each member
#resp2[0].dLExpansion.mailbox.each do |x|
#  puts x.name
#end

# class EmailAddressType < BaseEmailAddressType
  #attr_accessor :name
  #attr_accessor :emailAddress
  #attr_accessor :routingType
  #attr_accessor :mailboxType
  #attr_accessor :itemId
# Any of the above can be called as methods against dLExpansion.mailbox
# here's the basic structure of an expandDL response
# Response is an Array of ExpandDLExpansionType objects
# ExpandDLExpansionType is anArray of Mailbox objects
# Mailbox objects are made up of EmailAddressType objects
# EmailAddressType objects have the following attrs
# name, emailAddress, routingType, mailboxType, itemID
# mailboxType will let you know if the publicDL has members that are publicDLs themselves.
# if so, you'll have to iterate over THAT publicDL as well to get a fully recursed Distribution List entry