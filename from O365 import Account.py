from O365 import Account

credentials = ('ff91f935-c699-4f05-8844-0570bd88b673', ' ' )


account = Account(credentials)
if account.authenticate(scopes=['basic', 'message_all'], redirect_uri='https://login.microsoftonline.com/common/oauth2/nativeclient'):
    print('Authenticated!')
    account.connection.refresh_token()
    m = account.new_message(resource='jonas.quintiens@gmail.com')
    m.to.add('jonas.quintiens@gmail.com')
    m.subject = 'Testing!'
    m.body = "George Best quote: I've stopped drinking, but only while I'm asleep."
    m.send()