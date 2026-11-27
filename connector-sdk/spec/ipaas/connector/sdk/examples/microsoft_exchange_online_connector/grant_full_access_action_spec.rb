require 'spec_helper'

describe 'Microsoft Exchange Online Grant shared mailbox Full Access Action', :action, :microsoft_exchange_online do
  let(:action_template_id) { '01a1f2c3-0279-7d47-9565-a937427dd458' }

  def input(overrides = {})
    { mailbox: 'support@contoso.com', user: 'jane@contoso.com', auto_mapping: nil }.merge(overrides)
  end

  let(:lookup) { { Identity: 'support@contoso.com', User: 'jane@contoso.com' } }

  before do
    stub_cmdlet('Get-Mailbox', output: [{ RecipientTypeDetails: 'SharedMailbox' }],
                               with: { Identity: 'support@contoso.com' })
  end

  def grant(auto_mapping)
    { Identity: 'support@contoso.com', User: 'jane@contoso.com', AccessRights: ['FullAccess'],
      InheritanceType: 'All', AutoMapping: auto_mapping, }
  end

  describe 'when the user has no Full Access yet' do
    before do
      stub_cmdlet('Get-MailboxPermission', output: [], with: lookup)
      stub_cmdlet('Add-MailboxPermission', output: [{ AccessRights: ['FullAccess'] }])
    end

    it 'grants Full Access with auto-mapping and reports the change' do
      output = run_action(input)
      expect(cmdlet_request('Get-MailboxPermission', lookup)).to have_been_made
      expect(cmdlet_request('Add-MailboxPermission', grant(true))).to have_been_made
      expect(output).to eq('mailbox' => 'support@contoso.com', 'user' => 'jane@contoso.com',
                           'auto_mapping' => true, 'changed' => true)
    end

    it 'grants without auto-mapping when the step turns it off' do
      output = run_action(input(auto_mapping: false))
      expect(cmdlet_request('Add-MailboxPermission', grant(false))).to have_been_made
      expect(output['auto_mapping']).to be(false)
    end

    it 'sends every cmdlet to the tenant on the Exchange host with the Exchange token' do
      run_action(input)
      expect(a_request(:post, invoke_url).with(
               headers: { 'Authorization' => 'Bearer exchange-token', 'Content-Type' => 'application/json' },
             )).to have_been_made.times(3)
    end

    it 'types the access rights list for Exchange' do
      run_action(input)
      expect(a_request(:post, invoke_url).with do |request|
        parameters = JSON.parse(request.body).dig('CmdletInput', 'Parameters')
        parameters['AccessRights'] == ['FullAccess'] && parameters['AccessRights@odata.type'] == '#Collection(String)'
      end).to have_been_made
    end

    it 'strips the spaces around the mailbox and the user' do
      run_action(input(mailbox: '  support@contoso.com ', user: ' jane@contoso.com  '))
      expect(cmdlet_request('Add-MailboxPermission', grant(true))).to have_been_made
    end
  end

  describe 'when the user already has Full Access' do
    it 'changes nothing and reports no change' do
      stub_cmdlet('Get-MailboxPermission',
                  output: [{ AccessRights: ['FullAccess'], Deny: false, IsInherited: false, User: 'CONTOSO\\jane' }],
                  with: lookup)
      output = run_action(input)
      expect(cmdlet_request('Add-MailboxPermission')).not_to have_been_made
      expect(output['changed']).to be(false)
    end

    it 'reads access rights that Exchange wrote as one comma-separated string' do
      stub_cmdlet('Get-MailboxPermission',
                  output: [{ AccessRights: 'ReadPermission, FullAccess', Deny: 'False', IsInherited: 'False' }],
                  with: lookup)
      expect(run_action(input)['changed']).to be(false)
    end
  end

  describe 'permissions that do not count as Full Access' do
    before { stub_cmdlet('Add-MailboxPermission', output: []) }

    it 'grants when the only Full Access entry is a deny' do
      stub_cmdlet('Get-MailboxPermission', output: [{ AccessRights: ['FullAccess'], Deny: true, IsInherited: false }],
                                           with: lookup)
      expect(run_action(input)['changed']).to be(true)
    end

    it 'grants when the only Full Access entry is inherited' do
      stub_cmdlet('Get-MailboxPermission', output: [{ AccessRights: ['FullAccess'], Deny: false, IsInherited: true }],
                                           with: lookup)
      expect(run_action(input)['changed']).to be(true)
    end

    it 'grants when the user holds other rights only' do
      stub_cmdlet('Get-MailboxPermission', output: [{ AccessRights: ['ReadPermission'] }], with: lookup)
      expect(run_action(input)['changed']).to be(true)
    end
  end

  describe 'a mailbox that is not shared' do
    it 'refuses a user mailbox, so a step cannot open up a person\'s own mailbox' do
      stub_cmdlet('Get-Mailbox', output: [{ RecipientTypeDetails: 'UserMailbox' }],
                                 with: { Identity: 'support@contoso.com' })
      expect { run_action(input) }
        .to raise_error(IPaaS::Job::FailJob, "Mailbox 'support@contoso.com' is a UserMailbox, not a shared mailbox. " \
                                             'This step only opens up shared mailboxes.')
      expect(cmdlet_request('Get-MailboxPermission')).not_to have_been_made
    end

    it 'refuses a mailbox whose type Exchange does not report' do
      stub_cmdlet('Get-Mailbox', output: [{}], with: { Identity: 'support@contoso.com' })
      expect { run_action(input) }.to raise_error(IPaaS::Job::FailJob, /is a mailbox of unknown type, not a shared/)
    end
  end

  describe 'input checks' do
    it 'refuses a user that holds a routing character' do
      expect { run_action(input(user: 'contoso\\jane')) }
        .to raise_error(IPaaS::Job::FailJob, /User 'contoso\\jane' holds a slash/)
    end
  end
end
