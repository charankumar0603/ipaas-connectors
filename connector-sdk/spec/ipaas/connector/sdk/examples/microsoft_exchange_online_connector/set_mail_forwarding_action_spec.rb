require 'spec_helper'

describe 'Microsoft Exchange Online Set mail forwarding Action', :action, :microsoft_exchange_online do
  let(:action_template_id) { '01a1f2c3-028f-7dd1-8674-5b535937d270' }
  let(:manager_id) { '3f2a9c1e-7b4d-4e8a-9c61-2d5e8f0a7b13' }

  def input(overrides = {})
    { mailbox: 'leaver@contoso.com', target: 'manager@contoso.com', keep_copy: nil }.merge(overrides)
  end

  def stub_target(type)
    stub_cmdlet('Get-Recipient', output: [{ Name: manager_id, ExternalDirectoryObjectId: manager_id,
                                            PrimarySmtpAddress: 'manager@contoso.com', RecipientTypeDetails: type, }],
                                 with: { Identity: 'manager@contoso.com' })
  end

  def stub_mailbox(forwarding)
    stub_cmdlet('Get-Mailbox', output: [forwarding], with: { Identity: 'leaver@contoso.com' })
  end

  def forward(keep_copy)
    { Identity: 'leaver@contoso.com', ForwardingAddress: 'manager@contoso.com', ForwardingSmtpAddress: nil,
      DeliverToMailboxAndForward: keep_copy, }
  end

  describe 'to a mailbox in the organization' do
    before do
      stub_target('UserMailbox')
      stub_cmdlet('Set-Mailbox', output: [])
    end

    it 'sets the forwarding recipient, clears any SMTP forward, keeps a copy by default and reports the change' do
      stub_mailbox({ ForwardingAddress: nil, ForwardingSmtpAddress: nil, DeliverToMailboxAndForward: false })
      output = run_action(input)
      expect(cmdlet_request('Set-Mailbox', forward(true))).to have_been_made
      expect(output).to eq('mailbox' => 'leaver@contoso.com', 'target' => 'manager@contoso.com',
                           'keep_copy' => true, 'changed' => true)
    end

    it 'forwards without a copy when the step turns Keep a copy off' do
      stub_mailbox({ ForwardingAddress: manager_id, ForwardingSmtpAddress: nil, DeliverToMailboxAndForward: true })
      expect(run_action(input(keep_copy: false))).to include('keep_copy' => false, 'changed' => true)
      expect(cmdlet_request('Set-Mailbox', forward(false))).to have_been_made
    end

    it 'changes nothing when the mailbox already forwards to the target, named by its object ID' do
      stub_mailbox({ ForwardingAddress: manager_id, ForwardingSmtpAddress: nil, DeliverToMailboxAndForward: true })
      expect(run_action(input)['changed']).to be(false)
      expect(cmdlet_request('Set-Mailbox')).not_to have_been_made
    end

    it 'replaces an SMTP forward set beside it, which Exchange would ignore while it holds a recipient' do
      stub_mailbox({ ForwardingAddress: manager_id, ForwardingSmtpAddress: 'smtp:old@contoso.com',
                     DeliverToMailboxAndForward: true, })
      expect(run_action(input)['changed']).to be(true)
      expect(cmdlet_request('Set-Mailbox', forward(true))).to have_been_made
    end

    it 'replaces a forward to another recipient' do
      stub_mailbox({ ForwardingAddress: 'someone-else', ForwardingSmtpAddress: nil, DeliverToMailboxAndForward: true })
      expect(run_action(input)['changed']).to be(true)
    end
  end

  it 'forwards to a shared mailbox' do
    stub_target('SharedMailbox')
    stub_mailbox({ ForwardingAddress: nil, DeliverToMailboxAndForward: false })
    stub_cmdlet('Set-Mailbox', output: [])
    expect(run_action(input)['changed']).to be(true)
  end

  %w[MailContact MailUser MailUniversalDistributionGroup].each do |type|
    it "refuses a #{type}, whose mail can leave the organization" do
      stub_target(type)
      expect { run_action(input) }
        .to raise_error(IPaaS::Job::FailJob, "Target 'manager@contoso.com' is a #{type}. Mail is only forwarded to " \
                                             'a user or shared mailbox in the organization.')
      expect(cmdlet_request('Get-Mailbox')).not_to have_been_made
    end
  end

  it 'refuses an address that no recipient in the tenant holds, such as one in a relay domain' do
    stub_cmdlet_error('Get-Recipient', status: 404, exception: 'ManagementObjectNotFoundException',
                                       with: { Identity: 'x@partner.com' })
    expect { run_action(input(target: 'x@partner.com')) }
      .to raise_error(IPaaS::Job::FailJob, "Target 'x@partner.com' is not a recipient in this tenant. Mail is only " \
                                           'forwarded to a user or shared mailbox in the organization.')
  end
end
