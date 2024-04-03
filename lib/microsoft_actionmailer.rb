require "microsoft_actionmailer/version"
require 'microsoft_actionmailer/railtie' if defined?(Rails)
require 'microsoft_actionmailer/api'

require 'httparty'
require 'net/http'
require 'uri'

module MicrosoftActionmailer

  GRAPH_HOST = 'https://graph.microsoft.com'.freeze

  class DeliveryMethod
    include MicrosoftActionmailer::Api

    attr_reader :access_token,
                :delivery_options,
                :sender

    def initialize params
      @access_token = params[:authorization]
      @sender = params[:sender]
      @delivery_options = params[:delivery_options] || {}
    end

    def deliver! mail
      before_send = delivery_options[:before_send]
      if before_send && before_send.respond_to?(:call)
        before_send.call(mail, message)
      end

      body = mail.body.encoded
      body = mail.html_part.body.encoded if mail.html_part.present?

      res = ms_send_message(token: access_token, subject: mail.subject, content: body, recipients: mail.to, sender: sender, attachments: mail.attachments)

      after_send = delivery_options[:after_send]
      if after_send && after_send.respond_to?(:call)
        after_send.call(mail, message)
      end
    end
  end
end
