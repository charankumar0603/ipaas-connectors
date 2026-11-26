shared_context 'microsoft exchange online', :microsoft_exchange_online do
  let(:connector_id) { '01a1f2c3-0276-7a6d-88ab-3d9d74512014' }
  let(:tenant_id) { '9c46a35a-b7cd-42ef-9ce5-9ea1481d40f1' }
  let(:token_url) { "https://login.microsoftonline.com/#{tenant_id}/oauth2/v2.0/token" }
  let(:invoke_url) { "https://outlook.office365.com/adminapi/beta/#{tenant_id}/InvokeCommand" }
  let(:graph_base_url) { 'https://graph.microsoft.com/v1.0' }

  let(:outbound_connection_config) do
    {
      credentials: {
        tenant_id: tenant_id,
        client_id: 'client-abc',
        client_secret: make_secret_string('secret-def'),
      },
      cloud: 'Global',
    }
  end

  def cmdlet_of(request)
    JSON.parse(request.body).dig('CmdletInput', 'CmdletName')
  end

  # The parameters as the connector meant them, without the type annotations Exchange needs.
  def parameters_of(request)
    parameters = JSON.parse(request.body).dig('CmdletInput', 'Parameters') || {}
    parameters.reject { |key, _| key.include?('@') }
              .transform_values { |value| value.is_a?(Hash) ? value.except('@odata.type') : value }
  end

  # Exchange answers each cmdlet on its own, so every stub matches the cmdlet named in the body. A stub
  # registered later wins in WebMock, which lets an example override one cmdlet of a shared setup.
  def stub_cmdlet(cmdlet, output: [], status: 200, error: nil, headers: {})
    body = status == 200 ? { '@odata.context' => "#{invoke_url}/$metadata", value: output } : { error: error }
    stub_request(:post, invoke_url)
      .with { |request| cmdlet_of(request) == cmdlet }
      .to_return(status: status, body: body.to_json, headers: headers.merge('Content-Type' => 'application/json'))
  end

  # The error body Exchange Online sends for a failed cmdlet: the exception and its message in the first detail.
  def stub_cmdlet_error(cmdlet, exception: nil, status: 400, message: 'Exchange said no.', code: 'BadRequest')
    detail = exception ? "Ex6F9304|Microsoft.Exchange.Management.Tasks.#{exception}|#{message}" : message
    stub_cmdlet(cmdlet, status: status,
                        error: { code: code, message: 'Error executing cmdlet',
                                 details: [{ code: 'Context', target: '', message: detail }], })
  end

  def cmdlet_request(cmdlet, parameters = nil)
    a_request(:post, invoke_url).with do |request|
      cmdlet_of(request) == cmdlet && (parameters.nil? || parameters_of(request) == parameters.deep_stringify_keys)
    end
  end

  def expect_reschedule_after(seconds, message = nil, &step)
    Timecop.freeze do
      expect(&step).to raise_error(IPaaS::Job::RescheduleJob, message) do |error|
        expect(error.reschedule_after).to eq(seconds.seconds.from_now)
      end
    end
  end

  before(:each) do
    load_fixture('microsoft_exchange_online_connector')
    stub_request(:post, token_url)
      .with(body: hash_including('scope' => 'https://outlook.office365.com/.default'))
      .to_return(status: 200, body: { access_token: 'exchange-token', expires_in: 3599 }.to_json)
    stub_request(:post, token_url)
      .with(body: hash_including('scope' => 'https://graph.microsoft.com/.default'))
      .to_return(status: 200, body: { access_token: 'graph-token', expires_in: 3599 }.to_json)
  end
end
