module MicrosoftActionmailer
  module Api

    def ms_send_message(token:, subject:, content:, recipients:, sender:, attachments: [])

      attachment_list = []
      attachments.each do |attachment|
        data = { "@odata.type": "#microsoft.graph.fileAttachment",
                 "name": attachment.filename,
                 "contentType": attachment.content_type,
                 "contentBytes": Base64.encode64(attachment.body.raw_source)
               }
        attachment_list << data
      end


      send_message_url = "/v1.0/users/#{sender}/sendMail"
      req_method = 'post'
      query = {
        "message": {
          "subject": subject,
          "body":{
            "contentType": "HTML",
            "content": content
          },
          "attachments": attachment_list,
          "toRecipients": recipients.map{|address|
            {
              "emailAddress": {
                "address": address
              }
            }
          }
        }
      }

      response = make_api_call send_message_url, token, query, req_method
      raise response.parsed_response.to_s || "Request returned #{response.code}" unless response.code == 202
      response
    end

    def make_api_call(endpoint, token, params = nil, req_method)
      headers = {
        'Authorization'=> "Bearer #{token}",
        'Content-Type' => 'application/json'
      }

      query = params || {}
      if req_method == 'get'
        HTTParty.get "#{GRAPH_HOST}#{endpoint}",
                   headers: headers,
                   query: query
      elsif req_method == 'post'
        HTTParty.post "#{GRAPH_HOST}#{endpoint}",
                   headers: headers,
                   body: query.to_json
      end
    end
  end
end
