require 'spec_helper'

describe 'Microsoft Exchange Online Set automatic replies Action', :action, :microsoft_exchange_online do
  let(:action_template_id) { '01a1f2c3-0291-76fa-936f-5446fad43616' }
  let(:settings_url) { "#{graph_base_url}/users/leaver@contoso.com/mailboxSettings" }
  let(:reply_url) { "#{settings_url}/automaticRepliesSetting" }

  let(:disabled_setting) do
    { status: 'disabled', externalAudience: 'all', internalReplyMessage: 'Old internal.',
      externalReplyMessage: 'Old external.',
      scheduledStartDateTime: { dateTime: '2026-11-26T13:00:00.0000000', timeZone: 'UTC' },
      scheduledEndDateTime: { dateTime: '2026-11-27T13:00:00.0000000', timeZone: 'UTC' }, }
  end

  def input(overrides = {})
    { mailbox: 'leaver@contoso.com', status: 'Enabled', internal_message: 'I have left.',
      external_message: 'Please write to support@contoso.com.', external_audience: nil,
      scheduled_start: nil, scheduled_end: nil, }.merge(overrides)
  end

  def stub_current(setting = disabled_setting)
    stub_request(:get, reply_url)
      .to_return(status: 200, body: { '@odata.context': 'x' }.merge(setting).to_json,
                 headers: { 'Content-Type' => 'application/json' })
  end

  def stub_settings(status: 200, body: { automaticRepliesSetting: {} }.to_json, headers: {})
    stub_request(:patch, settings_url).to_return(status: status, body: body, headers: headers)
  end

  def patched(setting)
    a_request(:patch, settings_url).with(body: { automaticRepliesSetting: setting }.to_json,
                                         headers: { 'Authorization' => 'Bearer graph-token',
                                                    'Content-Type' => 'application/json', })
  end

  describe 'with the current setting read' do
    before do
      stub_current
      stub_settings
    end

    it 'turns replies on with both messages for all external senders, using a Graph token' do
      output = run_action(input)
      expect(patched(status: 'alwaysEnabled', internalReplyMessage: 'I have left.',
                     externalReplyMessage: 'Please write to support@contoso.com.', externalAudience: 'all'))
        .to have_been_made
      expect(output).to eq('mailbox' => 'leaver@contoso.com', 'status' => 'Enabled', 'external_audience' => 'All',
                           'scheduled_start' => nil, 'scheduled_end' => nil, 'changed' => true)
    end

    it 'keeps the messages the mailbox already has when the step gives none' do
      run_action(input(internal_message: nil, external_message: nil))
      expect(patched(status: 'alwaysEnabled', externalAudience: 'all')).to have_been_made
    end

    it 'limits the external audience when the step asks' do
      run_action(input(external_audience: 'Contacts only'))
      expect(a_request(:patch, settings_url)
        .with(body: hash_including('automaticRepliesSetting' => hash_including('externalAudience' => 'contactsOnly'))))
        .to have_been_made
    end

    it 'schedules replies between Start and End, in UTC' do
      output = run_action(input(status: 'Scheduled', scheduled_start: '2026-12-01T09:00:00+01:00',
                                scheduled_end: '2026-12-15T17:00:00+01:00'))
      expect(a_request(:patch, settings_url).with(body: hash_including('automaticRepliesSetting' => hash_including(
        'status' => 'scheduled',
        'scheduledStartDateTime' => { 'dateTime' => '2026-12-01T08:00:00', 'timeZone' => 'UTC' },
        'scheduledEndDateTime' => { 'dateTime' => '2026-12-15T16:00:00', 'timeZone' => 'UTC' },
      )))).to have_been_made
      expect(output['scheduled_start']).to be_present
      expect(output['scheduled_end']).to be_present
    end

    it 'reports no change for replies that already are off, and sends nothing' do
      output = run_action(input(status: 'Disabled'))
      expect(a_request(:patch, settings_url)).not_to have_been_made
      expect(output).to include('status' => 'Disabled', 'external_audience' => nil, 'changed' => false)
      expect(output['scheduled_start']).to be_nil
    end
  end

  it 'turns replies off without touching the messages' do
    stub_current(disabled_setting.merge(status: 'alwaysEnabled'))
    stub_settings
    expect(run_action(input(status: 'Disabled'))['changed']).to be(true)
    expect(patched(status: 'disabled')).to have_been_made
  end

  it 'reports no change when replies already are on this way' do
    stub_current(disabled_setting.merge(status: 'alwaysEnabled', internalReplyMessage: 'I have left.',
                                        externalReplyMessage: 'Please write to support@contoso.com.'))
    expect(run_action(input)['changed']).to be(false)
    expect(a_request(:patch, settings_url)).not_to have_been_made
  end

  it 'reports no change for the same schedule, which Graph answers with seven decimals' do
    stub_current(disabled_setting.merge(
                   status: 'scheduled', internalReplyMessage: 'I have left.',
                   externalReplyMessage: 'Please write to support@contoso.com.',
                   scheduledStartDateTime: { dateTime: '2026-12-01T08:00:00.0000000', timeZone: 'UTC' },
                   scheduledEndDateTime: { dateTime: '2026-12-15T16:00:00.0000000', timeZone: 'UTC' },
                 ))
    output = run_action(input(status: 'Scheduled', scheduled_start: '2026-12-01T09:00:00+01:00',
                              scheduled_end: '2026-12-15T17:00:00+01:00'))
    expect(output['changed']).to be(false)
  end

  it 'changes a schedule whose end moved' do
    stub_current(disabled_setting.merge(
                   status: 'scheduled', internalReplyMessage: 'I have left.',
                   externalReplyMessage: 'Please write to support@contoso.com.',
                   scheduledStartDateTime: { dateTime: '2026-12-01T08:00:00.0000000', timeZone: 'UTC' },
                   scheduledEndDateTime: 'not a time',
                 ))
    stub_settings
    output = run_action(input(status: 'Scheduled', scheduled_start: '2026-12-01T09:00:00+01:00',
                              scheduled_end: '2026-12-15T17:00:00+01:00'))
    expect(output['changed']).to be(true)
  end

  it 'refuses a schedule without an End' do
    expect { run_action(input(status: 'Scheduled', scheduled_start: '2026-12-01T09:00:00Z')) }
      .to raise_error(IPaaS::Job::FailJob, 'Scheduled replies need both Start and End.')
  end

  it 'refuses a schedule without a Start' do
    expect { run_action(input(status: 'Scheduled', scheduled_end: '2026-12-01T09:00:00Z')) }
      .to raise_error(IPaaS::Job::FailJob, 'Scheduled replies need both Start and End.')
  end

  it 'refuses a schedule that ends before it starts' do
    expect do
      run_action(input(status: 'Scheduled', scheduled_start: '2026-12-15T09:00:00Z',
                       scheduled_end: '2026-12-01T09:00:00Z'))
    end.to raise_error(IPaaS::Job::FailJob, 'End must be after Start.')
  end

  it 'names the permission when Graph refuses the read' do
    stub_request(:get, reply_url)
      .to_return(status: 403, body: { error: { code: 'ErrorAccessDenied', message: 'Access is denied.' } }.to_json)
    expect { run_action(input) }
      .to raise_error(IPaaS::Job::FailJob, 'Microsoft Graph returned HTTP 403 ErrorAccessDenied for mailbox ' \
                                           "'leaver@contoso.com': Access is denied. Check that the app holds " \
                                           'MailboxSettings.ReadWrite.')
  end

  describe 'failures of the change' do
    before { stub_current }

    it 'quotes the raw body of a failure that carries no Graph error' do
      stub_settings(status: 400, body: 'Bad Request')
      expect { run_action(input) }
        .to raise_error(IPaaS::Job::FailJob,
                        "Microsoft Graph returned HTTP 400  for mailbox 'leaver@contoso.com': Bad Request")
    end

    it 'says there was no body when Graph sent none' do
      stub_settings(status: 500, body: '')
      expect { run_action(input) }.to raise_error(IPaaS::Job::FailJob, /: no response body\z/)
    end

    it 'pauses the step when Graph throttles' do
      stub_settings(status: 429, body: '', headers: { 'Retry-After' => '12' })
      expect_reschedule_after(12) { run_action(input) }
    end

    it 'pauses the step when Graph answers 503' do
      stub_settings(status: 503, body: '')
      expect_reschedule_after(60) { run_action(input) }
    end

    [502, 504].each do |status|
      it "fails when the answer to the change was lost with #{status}, because it may have been applied" do
        stub_settings(status: status, body: '')
        expect { run_action(input) }
          .to raise_error(IPaaS::Job::FailJob, 'Microsoft Graph answered the automatic replies update with HTTP ' \
                                               "#{status}. The change may have been applied: check it in Exchange " \
                                               'before running the step again.')
      end
    end

    it 'fails when the change was not answered in time, because it may have been applied' do
      stub_request(:patch, settings_url).to_raise(Faraday::TimeoutError.new('Net::ReadTimeout'))
      expect { run_action(input) }
        .to raise_error(IPaaS::Job::FailJob, /\AMicrosoft Graph did not answer the automatic replies update in time/)
    end

    it 'pauses the step when the connection to Graph could not be made' do
      stub_request(:patch, settings_url).to_raise(Faraday::ConnectionFailed.new('Connection refused'))
      expect_reschedule_after(30, /\AMicrosoft Graph did not answer \(Connection refused\)/) { run_action(input) }
    end
  end

  it 'pauses the step when the read of the current setting is throttled' do
    stub_request(:get, reply_url).to_return(status: 429, body: '', headers: { 'Retry-After' => '5' })
    expect_reschedule_after(5) { run_action(input) }
  end

  it 'pauses the step when the read of the current setting is not answered in time' do
    stub_request(:get, reply_url).to_raise(Faraday::TimeoutError.new('Net::ReadTimeout'))
    expect_reschedule_after(30, /\AMicrosoft Graph did not answer/) { run_action(input) }
  end

  it 'pauses the step when Graph sign-in is busy' do
    stub_request(:post, token_url).to_return(status: 503, body: '')
    expect_reschedule_after(30, 'Sign-in at login.microsoftonline.com is unavailable (HTTP 503). ' \
                                'The step runs again shortly.') { run_action(input) }
  end
end
