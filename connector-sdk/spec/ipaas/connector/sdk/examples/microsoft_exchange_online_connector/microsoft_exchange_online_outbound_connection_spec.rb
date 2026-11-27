require 'spec_helper'

describe 'Microsoft Exchange Online Connection', :outbound_connection, :microsoft_exchange_online do
  # Bind the connector's helpers to this connection, as the runtime does before it runs a proc, so an inner
  # helpers.other_helper call and http_post resolve against this connection.
  let(:bound_helpers) { connector.outbound_connection.helpers_definition.copy_for(outbound_connection) }

  before(:each) do
    helpers_proxy = bound_helpers.for_proc
    outbound_connection.define_singleton_method(:helpers) { helpers_proxy }
  end

  def call_helper(name, ...)
    bound_helpers.registered_helper(name).execute(...)
  end

  describe 'cloud' do
    it 'reads the Global hosts by default' do
      outbound_connection.config[:cloud] = nil
      expect(call_helper(:cloud)).to eq(login: 'https://login.microsoftonline.com',
                                        graph: 'https://graph.microsoft.com',
                                        exchange: 'https://outlook.office365.com')
    end

    it 'pairs every cloud with the Exchange host the PowerShell module calls there' do
      hosts = MicrosoftExchangeOnlineConnector::CLOUDS.transform_values { |cloud| cloud[:exchange] }
      expect(hosts).to eq('Global' => 'https://outlook.office365.com',
                          'US Government GCC High' => 'https://outlook.office365.us',
                          'US Government DoD' => 'https://outlook-dod.office365.us',
                          'China' => 'https://partner.outlook.cn')
    end

    it 'refuses a cloud it does not know' do
      outbound_connection.config[:cloud] = 'Mars'
      expect { call_helper(:cloud) }.to raise_error(IPaaS::Job::FailJob, /Cloud 'Mars' is not one of Global/)
    end
  end

  describe 'tenant_id' do
    it 'refuses an empty tenant' do
      outbound_connection.config[:credentials][:tenant_id] = ' '
      expect { call_helper(:tenant_id) }.to raise_error(IPaaS::Job::FailJob, 'Tenant ID is empty on the connection.')
    end

    it 'refuses a sign-in placeholder' do
      outbound_connection.config[:credentials][:tenant_id] = 'Common'
      expect { call_helper(:tenant_id) }.to raise_error(IPaaS::Job::FailJob, /'Common' is a sign-in placeholder/)
    end
  end

  describe 'authorization_for' do
    it 'asks for a token for the given host and answers the Authorization header' do
      expect(call_helper(:authorization_for, 'https://outlook.office365.com')).to eq('Bearer exchange-token')
      expect(a_request(:post, token_url).with(body: hash_including(
        'client_id' => 'client-abc', 'client_secret' => 'secret-def', 'grant_type' => 'client_credentials',
        'scope' => 'https://outlook.office365.com/.default',
      ))).to have_been_made
    end

    it 'asks the China login host for a token for the China Exchange host' do
      outbound_connection.config[:cloud] = 'China'
      china_url = "https://login.chinacloudapi.cn/#{tenant_id}/oauth2/v2.0/token"
      stub_request(:post, china_url)
        .with(body: hash_including('scope' => 'https://partner.outlook.cn/.default'))
        .to_return(status: 200, body: { access_token: 'china-token', expires_in: 3599 }.to_json)
      expect(call_helper(:authorization_for, 'https://partner.outlook.cn')).to eq('Bearer china-token')
    end
  end

  describe 'invoke_command_url' do
    it 'names the tenant on the Exchange host of the cloud' do
      expect(call_helper(:invoke_command_url)).to eq(invoke_url)
      outbound_connection.config[:cloud] = 'US Government GCC High'
      expect(call_helper(:invoke_command_url))
        .to eq("https://outlook.office365.us/adminapi/beta/#{tenant_id}/InvokeCommand")
    end

    it 'escapes a tenant given as a domain name with unusual characters' do
      outbound_connection.config[:credentials][:tenant_id] = 'contoso#1.onmicrosoft.com'
      expect(call_helper(:invoke_command_url))
        .to eq('https://outlook.office365.com/adminapi/beta/contoso%231.onmicrosoft.com/InvokeCommand')
    end
  end

  describe 'escape_segment' do
    it 'escapes the characters that would end or reroute a URL path segment' do
      expect(call_helper(:escape_segment, 'a%b#c?d/e f')).to eq('a%25b%23c%3Fd%2Fe%20f')
    end
  end

  describe 'typed_parameters' do
    it 'types add and remove lists and string lists, and leaves plain values alone' do
      typed = call_helper(:typed_parameters, { Identity: 'bob', EmailAddresses: { add: ['smtp:bob2@contoso.com'] },
                                               AccessRights: ['FullAccess'], AutoMapping: true, })
      expect(typed).to eq(Identity: 'bob',
                          EmailAddresses: { add: ['smtp:bob2@contoso.com'],
                                            '@odata.type': '#Exchange.GenericHashTable', },
                          AccessRights: ['FullAccess'], 'AccessRights@odata.type': '#Collection(String)',
                          AutoMapping: true)
    end
  end

  describe 'exchange_cmdlet' do
    let(:parameters) { { Identity: 'support@contoso.com' } }

    let(:system_mailbox_guid) { 'bb558c35-97f1-4cb9-8ff7-d53741dc928c' }

    it 'posts the cmdlet to the tenant with the Exchange token and the routing header, and answers the values' do
      stub_cmdlet('Get-Mailbox', output: [{ Name: 'Support' }], with: parameters)
      expect(call_helper(:exchange_cmdlet, 'Get-Mailbox', parameters)).to eq([{ 'Name' => 'Support' }])
      expect(a_request(:post, invoke_url).with(
               headers: { 'Authorization' => 'Bearer exchange-token', 'Content-Type' => 'application/json',
                          'X-ResponseFormat' => 'json',
                          'X-AnchorMailbox' => "APP:SystemMailbox{#{system_mailbox_guid}}@#{tenant_id}", },
               body: { CmdletInput: { CmdletName: 'Get-Mailbox', Parameters: parameters } }.to_json,
             )).to have_been_made
    end

    it 'drops the type annotations Exchange puts on values, at every depth' do
      plan = { 'Name@odata.type' => '#String', 'Name' => 'P2' }
      stub_cmdlet('Get-Mailbox', output: [{ 'Name' => 'Support', 'ExchangeGuid@odata.type' => '#Guid',
                                            'ExchangeGuid' => 'e61c9754', 'Plan' => plan, }])
      expect(call_helper(:exchange_cmdlet, 'Get-Mailbox', parameters))
        .to eq([{ 'Name' => 'Support', 'ExchangeGuid' => 'e61c9754', 'Plan' => { 'Name' => 'P2' } }])
    end

    it 'answers an empty list when a cmdlet that changes something sends no values' do
      stub_request(:post, invoke_url).to_return(status: 200, body: '{}')
      expect(call_helper(:exchange_cmdlet, 'Set-Mailbox', parameters)).to eq([])
    end

    it 'answers an empty list for an empty body' do
      stub_request(:post, invoke_url).to_return(status: 200, body: '')
      expect(call_helper(:exchange_cmdlet, 'Set-Mailbox', parameters)).to eq([])
    end

    it 'fails on a success that is not JSON, so a proxy page is not taken for an empty answer' do
      stub_request(:post, invoke_url).to_return(status: 200, body: '<html>Sign in</html>')
      expect { call_helper(:exchange_cmdlet, 'Get-Recipient', parameters) }
        .to raise_error(IPaaS::Job::FailJob, /not valid JSON/)
    end

    it 'skips the Graph token, which Exchange does not accept' do
      stub_cmdlet('Get-Mailbox', output: [])
      call_helper(:exchange_cmdlet, 'Get-Mailbox', parameters)
      expect(a_request(:post, token_url).with(body: hash_including('scope' => 'https://graph.microsoft.com/.default')))
        .not_to have_been_made
    end

    it 'fails when Exchange answers success with values that are not a list' do
      stub_request(:post, invoke_url).to_return(status: 200, body: { value: 'done' }.to_json)
      expect { call_helper(:exchange_cmdlet, 'Get-Mailbox', parameters) }
        .to raise_error(IPaaS::Job::FailJob, 'Exchange Online answered Get-Mailbox without a value list.')
    end

    it 'fails when Exchange answers success with a JSON list instead of an object' do
      stub_request(:post, invoke_url).to_return(status: 200, body: '"done"')
      expect { call_helper(:exchange_cmdlet, 'Get-Mailbox', parameters) }
        .to raise_error(IPaaS::Job::FailJob, 'Exchange Online answered Get-Mailbox without a value list.')
    end

    it 'answers nil for the exception the caller tolerates, matched by its short name only' do
      stub_cmdlet_error('Get-Recipient', status: 404, exception: 'ManagementObjectNotFoundException')
      expect(call_helper(:exchange_cmdlet, 'Get-Recipient', parameters,
                         tolerate: 'ManagementObjectNotFoundException')).to be_nil
      expect { call_helper(:exchange_cmdlet, 'Get-Recipient', parameters, tolerate: 'ObjectNotFoundException') }
        .to raise_error(IPaaS::Job::FailJob, /ManagementObjectNotFoundException/)
    end

    it 'quotes the exception and message Exchange put in the error details' do
      stub_cmdlet_error('Get-Mailbox', status: 404, code: 'NotFound', exception: 'ManagementObjectNotFoundException',
                                       message: "The operation couldn't be performed because object couldn't be found.")
      expect { call_helper(:exchange_cmdlet, 'Get-Mailbox', parameters) }
        .to raise_error(IPaaS::Job::FailJob, 'Get-Mailbox failed with HTTP 404 (ManagementObjectNotFoundException): ' \
                                             "The operation couldn't be performed because object couldn't be found.")
    end

    it 'quotes the detail as the message when it names no exception' do
      stub_cmdlet_error('Get-Mailbox', status: 400, message: 'A parameter cannot be found.')
      expect { call_helper(:exchange_cmdlet, 'Get-Mailbox', parameters) }
        .to raise_error(IPaaS::Job::FailJob, 'Get-Mailbox failed with HTTP 400: A parameter cannot be found.')
    end

    it 'skips details that are not objects' do
      body = { error: { code: 'BadRequest', message: 'Error executing cmdlet',
                        details: ['odd', { message: 'Ex1|Microsoft.Exchange.X.RecipientTaskException|No.' }], } }
      stub_request(:post, invoke_url).to_return(status: 400, body: body.to_json)
      expect { call_helper(:exchange_cmdlet, 'Get-Mailbox', parameters) }
        .to raise_error(IPaaS::Job::FailJob, 'Get-Mailbox failed with HTTP 400 (RecipientTaskException): No.')
    end

    it 'quotes the raw body when Exchange sent no error object' do
      stub_request(:post, invoke_url).to_return(status: 500, body: 'Internal Server Error')
      expect { call_helper(:exchange_cmdlet, 'Get-AcceptedDomain', {}) }
        .to raise_error(IPaaS::Job::FailJob, 'Get-AcceptedDomain failed with HTTP 500: Internal Server Error')
    end

    it 'names the permission, consent, tenant, cloud and token lifetime when Exchange refuses the token' do
      stub_request(:post, invoke_url)
        .to_return(status: 401, body: { error: { code: 'Unauthorized', message: 'Token is not valid.' } }.to_json)
      expect { call_helper(:exchange_cmdlet, 'Get-AcceptedDomain', {}) }
        .to raise_error(IPaaS::Job::FailJob, "Exchange Online refused the app's token for Get-AcceptedDomain: " \
                                             'Token is not valid. Check that the app holds Exchange.ManageAsApp ' \
                                             '(Office 365 Exchange Online) with admin consent, and that Tenant ID ' \
                                             "and Cloud match it. A change to the app's permissions or roles " \
                                             'reaches the connector within an hour, when its cached token expires.')
    end

    it 'names the Exchange role when Exchange refuses the cmdlet' do
      stub_cmdlet_error('Get-AcceptedDomain', status: 403, code: 'Forbidden', message: 'Access denied.')
      expect { call_helper(:exchange_cmdlet, 'Get-AcceptedDomain', {}) }
        .to raise_error(IPaaS::Job::FailJob, 'Exchange Online refused Get-AcceptedDomain: Access denied. Check ' \
                                             'that the app holds Exchange.ManageAsApp (Office 365 Exchange ' \
                                             "Online; the Admin API's Exchange.AdminAPI.ManageAsApp does not do) " \
                                             "with admin consent, and that the app's service principal holds an " \
                                             'Exchange role that includes this cmdlet. A change to the app\'s ' \
                                             'permissions or roles reaches the connector within an hour, when its ' \
                                             'cached token expires.')
    end

    it 'says Exchange gave no reason when it refuses with an empty body, as it does without Exchange.ManageAsApp' do
      stub_request(:post, invoke_url).to_return(status: 403, body: '')
      expect { call_helper(:exchange_cmdlet, 'Get-AcceptedDomain', {}) }
        .to raise_error(IPaaS::Job::FailJob, /\AExchange Online refused Get-AcceptedDomain: Exchange gave no reason\. /)
    end

    it 'pauses the step when Exchange throttles the tenant' do
      stub_request(:post, invoke_url).to_return(status: 429, body: '', headers: { 'Retry-After' => '30' })
      expect_reschedule_after(30) { call_helper(:exchange_cmdlet, 'Set-Mailbox', parameters) }
    end

    [502, 503, 504].each do |status|
      it "pauses a read when Exchange answers #{status}" do
        stub_request(:post, invoke_url).to_return(status: status, body: '')
        expect_reschedule_after(60) { call_helper(:exchange_cmdlet, 'Get-Mailbox', parameters) }
      end
    end

    it 'pauses a change when Exchange answers 503, which it sends before running the cmdlet' do
      stub_request(:post, invoke_url).to_return(status: 503, body: '')
      expect_reschedule_after(60) { call_helper(:exchange_cmdlet, 'Set-Mailbox', parameters) }
    end

    [502, 504].each do |status|
      it "fails a change whose answer was lost with #{status}, because it may have been applied" do
        stub_request(:post, invoke_url).to_return(status: status, body: '')
        expect { call_helper(:exchange_cmdlet, 'Set-Mailbox', parameters) }
          .to raise_error(IPaaS::Job::FailJob, "Exchange Online answered Set-Mailbox with HTTP #{status}. The " \
                                               'change may have been applied: check it in Exchange before running ' \
                                               'the step again.')
      end
    end

    it 'pauses a read whose answer did not arrive in time' do
      stub_request(:post, invoke_url).to_raise(Faraday::TimeoutError.new('Net::ReadTimeout'))
      expect_reschedule_after(30, /\AExchange Online did not answer \(.+\)\. The step runs again shortly\.\z/) do
        call_helper(:exchange_cmdlet, 'Get-Mailbox', parameters)
      end
    end

    it 'fails a change whose answer did not arrive in time, because it may have been applied' do
      stub_request(:post, invoke_url).to_raise(Faraday::TimeoutError.new('Net::ReadTimeout'))
      expect { call_helper(:exchange_cmdlet, 'Add-MailboxPermission', parameters) }
        .to raise_error(IPaaS::Job::FailJob, /\AExchange Online did not answer Add-MailboxPermission in time \(.+\)\. /)
    end

    it 'pauses even a change when the connection to Exchange could not be made' do
      stub_request(:post, invoke_url).to_raise(Faraday::ConnectionFailed.new('Connection refused'))
      expect_reschedule_after(30, 'Exchange Online did not answer (Connection refused). ' \
                                  'The step runs again shortly.') do
        call_helper(:exchange_cmdlet, 'Set-Mailbox', parameters)
      end
    end

    [429, 500, 503].each do |status|
      it "pauses the step when sign-in answers #{status}" do
        stub_request(:post, token_url).to_return(status: status, body: '')
        expect_reschedule_after(30, "Sign-in at login.microsoftonline.com is unavailable (HTTP #{status}). " \
                                    'The step runs again shortly.') do
          call_helper(:exchange_cmdlet, 'Set-Mailbox', parameters)
        end
        expect(a_request(:post, invoke_url)).not_to have_been_made
      end
    end

    it 'pauses the step and names the sign-in host when sign-in does not answer' do
      stub_request(:post, token_url).to_timeout
      expect_reschedule_after(30, /\ASign-in at login\.microsoftonline\.com did not answer/) do
        call_helper(:exchange_cmdlet, 'Set-Mailbox', parameters)
      end
    end

    it 'fails the step when sign-in refuses the credentials' do
      stub_request(:post, token_url)
        .to_return(status: 401, body: { error: 'invalid_client', error_description: 'Bad secret.' }.to_json)
      expect { call_helper(:exchange_cmdlet, 'Get-Mailbox', parameters) }
        .to raise_error(IPaaS::Job::Outbound::CustomerCredentialsError, /invalid_client: Bad secret\./)
    end

    it 'keeps an unexpected error as it is' do
      stub_request(:post, invoke_url).to_raise(ArgumentError.new('bad argument'))
      expect { call_helper(:exchange_cmdlet, 'Get-Mailbox', parameters) }.to raise_error(ArgumentError, 'bad argument')
    end
  end

  describe 'authenticate' do
    let(:request) { Faraday::Request.create(:get) { |built| built.headers = {} } }

    it 'puts a Graph token on the request' do
      outbound_connection.authenticate_request(request)
      expect(request.headers['Authorization']).to eq('Bearer graph-token')
    end
  end

  describe 'config_tester' do
    it 'reads the accepted domains and says Exchange answered' do
      stub_cmdlet('Get-AcceptedDomain', output: [{ DomainName: 'contoso.com' }])
      expect(outbound_connection.config_tester)
        .to eq(status: :success, message: "Connection successful. Exchange Online answered for tenant #{tenant_id}.")
      expect(cmdlet_request('Get-AcceptedDomain', {})).to have_been_made
    end

    it 'reports failed with the reason when Exchange refuses' do
      stub_cmdlet_error('Get-AcceptedDomain', status: 401, message: 'Unauthorized.')
      expect(outbound_connection.config_tester[:status]).to eq(:failed)
      expect(outbound_connection.config_tester[:message]).to start_with("Exchange Online refused the app's token")
    end

    it 'says the test could not run, rather than that it runs again, when Exchange is busy' do
      stub_request(:post, invoke_url).to_return(status: 503, body: '')
      result = outbound_connection.config_tester
      expect(result[:status]).to eq(:error)
      expect(result[:message]).to start_with('Exchange Online is busy or did not answer, so the connection could ' \
                                             'not be tested. Try again in a minute.')
    end
  end
end
