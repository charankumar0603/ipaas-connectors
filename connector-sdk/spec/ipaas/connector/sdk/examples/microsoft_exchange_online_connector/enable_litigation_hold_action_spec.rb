require 'spec_helper'

describe 'Microsoft Exchange Online Enable litigation hold Action', :action, :microsoft_exchange_online do
  let(:action_template_id) { '01a1f2c3-029b-71e5-95cc-bb9dabb2aafa' }

  def input(overrides = {})
    { mailbox: 'leaver@contoso.com', duration_days: nil, owner: nil, comment: nil }.merge(overrides)
  end

  def hold_call_with(parameter)
    a_request(:post, invoke_url).with do |request|
      cmdlet_of(request) == 'Set-Mailbox' && parameters_of(request).key?(parameter)
    end
  end

  it 'places an unlimited hold by default and reports the change' do
    stub_cmdlet('Get-Mailbox', output: [{ LitigationHoldEnabled: false, LitigationHoldDuration: 'Unlimited' }],
                               with: { Identity: 'leaver@contoso.com' })
    stub_cmdlet('Set-Mailbox', output: [])
    output = run_action(input)
    expect(cmdlet_request('Set-Mailbox', { Identity: 'leaver@contoso.com', LitigationHoldEnabled: true,
                                           LitigationHoldDuration: 'Unlimited', })).to have_been_made
    expect(output).to eq('mailbox' => 'leaver@contoso.com', 'duration' => 'Unlimited', 'changed' => true)
  end

  it 'places a hold for a number of days with its owner and comment' do
    stub_cmdlet('Get-Mailbox', output: [{ LitigationHoldEnabled: false }], with: { Identity: 'leaver@contoso.com' })
    stub_cmdlet('Set-Mailbox', output: [])
    output = run_action(input(duration_days: 365, owner: 'Legal', comment: 'Case 2026-114'))
    expect(cmdlet_request('Set-Mailbox', { Identity: 'leaver@contoso.com', LitigationHoldEnabled: true,
                                           LitigationHoldDuration: 365, RetentionComment: 'Case 2026-114', }))
      .to have_been_made
    expect(cmdlet_request('Set-Mailbox', { Identity: 'leaver@contoso.com', LitigationHoldOwner: 'Legal' }))
      .to have_been_made
    expect(output['duration']).to eq('365')
  end

  it 'sets no owner when none is given, so Exchange records the app' do
    stub_cmdlet('Get-Mailbox', output: [{ LitigationHoldEnabled: false }], with: { Identity: 'leaver@contoso.com' })
    stub_cmdlet('Set-Mailbox', output: [])
    run_action(input(owner: '  '))
    expect(a_request(:post, invoke_url)
      .with { |request| parameters_of(request).key?('LitigationHoldOwner') }).not_to have_been_made
  end

  it 'changes nothing when the same hold is already in place' do
    stub_cmdlet('Get-Mailbox', output: [{ LitigationHoldEnabled: true, LitigationHoldDuration: '365.00:00:00' }],
                               with: { Identity: 'leaver@contoso.com' })
    expect(run_action(input(duration_days: 365))['changed']).to be(false)
    expect(cmdlet_request('Set-Mailbox')).not_to have_been_made
  end

  it 'treats an existing unlimited hold as the same as an unlimited request' do
    stub_cmdlet('Get-Mailbox', output: [{ LitigationHoldEnabled: true, LitigationHoldDuration: 'unlimited' }],
                               with: { Identity: 'leaver@contoso.com' })
    expect(run_action(input)['changed']).to be(false)
  end

  it 'changes a hold with another duration' do
    stub_cmdlet('Get-Mailbox', output: [{ LitigationHoldEnabled: true, LitigationHoldDuration: 'Unlimited' }],
                               with: { Identity: 'leaver@contoso.com' })
    stub_cmdlet('Set-Mailbox', output: [])
    expect(run_action(input(duration_days: 30))['changed']).to be(true)
  end

  describe 'a hold already in place' do
    let(:hold) do
      { LitigationHoldEnabled: true, LitigationHoldDuration: 'Unlimited', LitigationHoldOwner: 'Legal',
        RetentionComment: 'Case 7', }
    end

    before { stub_cmdlet('Set-Mailbox', output: []) }

    it 'changes nothing when the owner and comment match too, whatever the case of the owner' do
      stub_cmdlet('Get-Mailbox', output: [hold], with: { Identity: 'leaver@contoso.com' })
      expect(run_action(input(owner: 'legal', comment: 'Case 7'))['changed']).to be(false)
      expect(cmdlet_request('Set-Mailbox')).not_to have_been_made
    end

    it 'sets the owner a lost second call left as the app, without switching the hold again' do
      stub_cmdlet('Get-Mailbox', output: [hold.merge(LitigationHoldOwner: 'SystemMailbox{bb558c35}@contoso.com')],
                                 with: { Identity: 'leaver@contoso.com' })
      expect(run_action(input(owner: 'Legal', comment: 'Case 7'))['changed']).to be(true)
      expect(cmdlet_request('Set-Mailbox', { Identity: 'leaver@contoso.com', LitigationHoldOwner: 'Legal' }))
        .to have_been_made
      expect(hold_call_with('LitigationHoldEnabled')).not_to have_been_made
    end

    it 'writes a new comment and leaves the owner alone' do
      stub_cmdlet('Get-Mailbox', output: [hold], with: { Identity: 'leaver@contoso.com' })
      expect(run_action(input(owner: 'Legal', comment: 'Case 8'))['changed']).to be(true)
      expect(cmdlet_request('Set-Mailbox', { Identity: 'leaver@contoso.com', LitigationHoldEnabled: true,
                                             LitigationHoldDuration: 'Unlimited', RetentionComment: 'Case 8', }))
        .to have_been_made
      expect(hold_call_with('LitigationHoldOwner')).not_to have_been_made
    end

    it 'ignores the owner and comment the step does not name' do
      stub_cmdlet('Get-Mailbox', output: [hold], with: { Identity: 'leaver@contoso.com' })
      expect(run_action(input)['changed']).to be(false)
    end
  end

  it 'fails with the Exchange reason when the mailbox lacks the licence for a hold' do
    stub_cmdlet('Get-Mailbox', output: [{ LitigationHoldEnabled: false }], with: { Identity: 'leaver@contoso.com' })
    stub_cmdlet_error('Set-Mailbox', exception: 'RecipientTaskException',
                                     message: 'The mailbox needs an Exchange Online Plan 2 license.')
    expect { run_action(input) }.to raise_error(IPaaS::Job::FailJob, /needs an Exchange Online Plan 2 license/)
  end
end
