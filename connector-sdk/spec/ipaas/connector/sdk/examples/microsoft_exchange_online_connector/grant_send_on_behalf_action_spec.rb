require 'spec_helper'

describe 'Microsoft Exchange Online Grant Send on Behalf Action', :action, :microsoft_exchange_online do
  let(:action_template_id) { '01a1f2c3-0283-7a28-a20d-7151e9003538' }

  def input(overrides = {})
    { mailbox: 'ceo@contoso.com', user: 'jane@contoso.com' }.merge(overrides)
  end

  let(:jane) do
    { Name: 'jane.doe', Identity: 'jane-identity', Alias: 'jane', DisplayName: 'Jane Doe',
      PrimarySmtpAddress: 'jane.doe@contoso.com', ExternalDirectoryObjectId: '7f2b0c3e-1111-4c1a-9d3b-2f0a9e1c5b7d', }
  end

  before { stub_cmdlet('Get-Recipient', output: [jane], with: { Identity: 'jane@contoso.com' }) }

  it 'adds the user to the Send on Behalf list and keeps the delegates already on it' do
    stub_cmdlet('Get-Mailbox', output: [{ GrantSendOnBehalfTo: ['Blake Lee'] }], with: { Identity: 'ceo@contoso.com' })
    stub_cmdlet('Set-Mailbox', output: [])
    output = run_action(input)
    expect(cmdlet_request('Set-Mailbox', { Identity: 'ceo@contoso.com',
                                           GrantSendOnBehalfTo: { add: ['jane@contoso.com'] }, })).to have_been_made
    expect(output).to eq('mailbox' => 'ceo@contoso.com', 'user' => 'jane@contoso.com', 'changed' => true)
  end

  it 'adds the user when the mailbox has no delegates at all' do
    stub_cmdlet('Get-Mailbox', output: [{ GrantSendOnBehalfTo: nil }], with: { Identity: 'ceo@contoso.com' })
    stub_cmdlet('Set-Mailbox', output: [])
    expect(run_action(input)['changed']).to be(true)
  end

  it 'changes nothing when Exchange lists the user by object id, as one value rather than a list' do
    stub_cmdlet('Get-Mailbox', output: [{ GrantSendOnBehalfTo: '7f2b0c3e-1111-4c1a-9d3b-2f0a9e1c5b7d' }],
                               with: { Identity: 'ceo@contoso.com' })
    expect(run_action(input)['changed']).to be(false)
    expect(cmdlet_request('Set-Mailbox')).not_to have_been_made
  end

  [['its name', 'JANE.DOE'], ['its identity', 'jane-identity'], ['its primary address', 'Jane.Doe@contoso.com'],
   ['the address the step was given', 'JANE@contoso.com'],].each do |label, entry|
    it "changes nothing when Exchange lists the user by #{label}" do
      stub_cmdlet('Get-Mailbox', output: [{ GrantSendOnBehalfTo: [entry] }], with: { Identity: 'ceo@contoso.com' })
      expect(run_action(input)['changed']).to be(false)
    end
  end

  it 'does not take another person with the same display name or alias for the user' do
    stub_cmdlet('Get-Mailbox', output: [{ GrantSendOnBehalfTo: ['Jane Doe', 'jane'] }],
                               with: { Identity: 'ceo@contoso.com' })
    stub_cmdlet('Set-Mailbox', output: [])
    expect(run_action(input)['changed']).to be(true)
  end

  it 'fails when the mailbox does not exist' do
    stub_cmdlet('Get-Mailbox', output: [], with: { Identity: 'ceo@contoso.com' })
    expect { run_action(input) }
      .to raise_error(IPaaS::Job::FailJob, "Mailbox 'ceo@contoso.com' was not found in Exchange Online.")
  end

  it 'fails when the user does not exist' do
    stub_cmdlet('Get-Mailbox', output: [{ GrantSendOnBehalfTo: [] }], with: { Identity: 'ceo@contoso.com' })
    stub_cmdlet('Get-Recipient', output: [], with: { Identity: 'jane@contoso.com' })
    expect { run_action(input) }
      .to raise_error(IPaaS::Job::FailJob, "User 'jane@contoso.com' was not found in Exchange Online.")
  end
end
