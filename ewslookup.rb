require 'rubygems'
gem 'soap4r'
require 'defaultDriver.rb'
# EWS - Exchange Web Services interface
# Copyright (C) 2009  John E. Vincent <rubystuff@lusis.org>.

# This program is copyrighted free software by John E. Vincent.  You can
# redistribute it and/or modify it under the same terms of Ruby's license;
# either the dual license version in 2003, or any later version.
class EWSLookup
# Set up the parameters for connecting to an Exchange Server 2007 via Exchange Web Services
# EXAMPLE 1
#   EWS.new('bob','password','http://my.exchange.com/EWS/Exchange.asmx')
#
# ARGS
#   user: username
#   pass: password
#   endpoint: URl to Exchange.asmx (i.e. https://my.mail.server/ews/Exchange.asmx)
#
# RETURNS
#   EWSLookup instance. Use #gal_lookup or #dl_lookup to access the GAL and Distribution Lists respectively
# 
#
  def initialize(user,pass,endpoint)
    @user = user
    @pass = pass
    @endpoint = endpoint
  end
  
  def connect
    namespace = 'http://schemas.microsoft.com/exchange/services/2006/messages'
    ews = ExchangeServicePortType.new(@endpoint)
    ews.options["protocol.http.auth.ntlm"] = [@endpoint, @user, @pass]
    ews.options["protocol.http.ssl_config.verify_mode"] = nil 
    ews.wiredump_file_base = "soap-log.txt" if OPTIONS[:debug]
    return ews
  end
  
  def gal_lookup(query)
    resarr = Array.new
    ews = self.connect
    n = ResolveNamesType.new(nil, query)
    n.xmlattr_ReturnFullContactData=true
    r = ews.resolveNames(n).responseMessages.resolveNamesResponseMessage
    if r[0].xmlattr_ResponseClass == 'Error'
      return r[0].messageText
    end    
    r[0].resolutionSet.resolution.each do |entry|
	resarr << [entry.mailbox.name,entry.mailbox.emailAddress,entry.mailbox.mailboxType] #if entry.mailbox.mailboxType == 'Mailbox'
    end
    return resarr
  end
  
  def dl_lookup(query)
    resarr = Array.new
    ews = self.connect
    e = EmailAddressType.new('', query, 'SMTP')
    l = ExpandDLType.new(e)
    r = ews.expandDL(l).responseMessages.expandDLResponseMessage
    if r[0].xmlattr_ResponseClass == 'Error'
      return r[0].messageText
    end
    r[0].dLExpansion.mailbox.each do |member|
	resarr << [member.name,member.emailAddress]
    end
    return resarr
  end
end
