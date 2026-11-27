require 'spec_helper'

describe 'Microsoft Exchange Online Grant calendar delegate access Action', :action, :microsoft_exchange_online do
  let(:action_template_id) { '01a1f2c3-0296-72c4-a69e-0d7cfcaed557' }
  let(:calendar) { 'ceo@contoso.com:\\Kalender' }

  def input(overrides = {})
    { mailbox: 'ceo@contoso.com', user: 'jane@contoso.com', access_rights: 'Reviewer',
      sharing_permission: nil, }.merge(overrides)
  end

  let(:lookup) { { Identity: calendar, User: 'jane@contoso.com' } }

  def stub_no_entry
    stub_cmdlet_error('Get-MailboxFolderPermission', status: 404, code: 'NotFound',
                                                     exception: 'UserNotFoundInPermissionEntryException',
                                                     with: lookup)
  end

  def stub_entry(entry)
    stub_cmdlet('Get-MailboxFolderPermission', output: [entry], with: lookup)
  end

  before do
    stub_cmdlet('Get-MailboxFolderStatistics', output: [{ Name: 'Kalender', FolderType: 'Calendar' },
                                                        { Name: 'Geburtstage', FolderType: 'Birthdays' },],
                                               with: { Identity: 'ceo@contoso.com', FolderScope: 'Calendar' })
  end

  it 'finds the calendar by type, whatever its name, and adds the user' do
    stub_no_entry
    stub_cmdlet('Add-MailboxFolderPermission', output: [])
    output = run_action(input)
    expect(cmdlet_request('Get-MailboxFolderStatistics', { Identity: 'ceo@contoso.com', FolderScope: 'Calendar' }))
      .to have_been_made
    expect(cmdlet_request('Add-MailboxFolderPermission',
                          { Identity: calendar, User: 'jane@contoso.com',
                            AccessRights: ['Reviewer'], })).to have_been_made
    expect(output).to eq('mailbox' => 'ceo@contoso.com', 'user' => 'jane@contoso.com', 'calendar' => calendar,
                         'access_rights' => 'Reviewer', 'sharing_permission' => 'None', 'changed' => true)
  end

  it 'makes the user a delegate who sees private items when asked' do
    stub_no_entry
    stub_cmdlet('Add-MailboxFolderPermission', output: [])
    run_action(input(access_rights: 'Editor', sharing_permission: 'Delegate and private items'))
    expect(cmdlet_request('Add-MailboxFolderPermission',
                          { Identity: calendar, User: 'jane@contoso.com', AccessRights: ['Editor'],
                            SharingPermissionFlags: 'Delegate,CanViewPrivateItems', })).to have_been_made
  end

  it 'changes the rights of a user who already has other rights, saying the sharing flags too' do
    stub_entry({ AccessRights: ['AvailabilityOnly'] })
    stub_cmdlet('Set-MailboxFolderPermission', output: [])
    expect(run_action(input)['changed']).to be(true)
    expect(cmdlet_request('Set-MailboxFolderPermission',
                          { Identity: calendar, User: 'jane@contoso.com', AccessRights: ['Reviewer'],
                            SharingPermissionFlags: 'None', })).to have_been_made
    expect(cmdlet_request('Add-MailboxFolderPermission')).not_to have_been_made
  end

  it 'takes delegate status away when the step asks for None, which Exchange keeps unless told' do
    stub_entry({ AccessRights: ['Editor'], SharingPermissionFlags: 'Delegate' })
    stub_cmdlet('Set-MailboxFolderPermission', output: [])
    expect(run_action(input(access_rights: 'Editor'))['changed']).to be(true)
    expect(cmdlet_request('Set-MailboxFolderPermission',
                          { Identity: calendar, User: 'jane@contoso.com', AccessRights: ['Editor'],
                            SharingPermissionFlags: 'None', })).to have_been_made
  end

  it 'changes nothing when the user already has exactly these rights' do
    stub_entry({ AccessRights: ['Reviewer'], SharingPermissionFlags: nil })
    expect(run_action(input)['changed']).to be(false)
    expect(cmdlet_request('Set-MailboxFolderPermission')).not_to have_been_made
  end

  it 'compares delegate flags whatever order Exchange lists them in' do
    stub_cmdlet('Get-MailboxFolderPermission',
                output: [{ AccessRights: ['Editor'], SharingPermissionFlags: 'CanViewPrivateItems, Delegate' }],
                with: lookup)
    expect(run_action(input(access_rights: 'Editor', sharing_permission: 'Delegate and private items'))['changed'])
      .to be(false)
  end

  it 'refuses a delegate sharing permission without the Editor right' do
    expect { run_action(input(sharing_permission: 'Delegate')) }
      .to raise_error(IPaaS::Job::FailJob, 'Sharing permission Delegate needs the Editor access right, not Reviewer.')
  end

  it 'fails when the mailbox has no calendar folder' do
    stub_cmdlet('Get-MailboxFolderStatistics', output: [{ Name: 'Inbox', FolderType: 'Inbox' }],
                                               with: { Identity: 'ceo@contoso.com', FolderScope: 'Calendar' })
    expect do
      run_action(input)
    end.to raise_error(IPaaS::Job::FailJob, "Mailbox 'ceo@contoso.com' has no calendar folder.")
  end

  it 'fails when Exchange cannot read the folder permissions' do
    stub_cmdlet_error('Get-MailboxFolderPermission', status: 500, code: 'InternalServerError',
                                                     message: 'Server error.', with: lookup)
    expect { run_action(input) }
      .to raise_error(IPaaS::Job::FailJob, 'Get-MailboxFolderPermission failed with HTTP 500: Server error.')
  end
end
