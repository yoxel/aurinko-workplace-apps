const clientIdName = 'clientId';
const baseUrl = 'https://addin-v2.aurinko.io';
let params = new URLSearchParams(window.location.search);
let clientId = params.get(clientIdName);
let onAuthClick;

Office.initialize = function (reason) {
    console.log('Office.initialize started (' + reason + ')');

    let hostMailbox = {exchangeToken: '', email: ''};

    Office.context.mailbox.getUserIdentityTokenAsync(function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            console.error(`Token retrieval failed with message: ${result.error.message}`);
        } else {

            hostMailbox.exchangeToken = result.value;

            let authHeaders = {
                'X-Aurinko-ClientId': clientId,
                'X-Aurinko-AuthType': 'exchangeIdToken',
                'authorization': 'Bearer ' + hostMailbox.exchangeToken,
                'content-type': 'application/json'
            };

            onAuthClick = function () {

                request('POST', baseUrl + '/v1/auth/prepare', authHeaders, function (resp) {

                    let token = JSON.parse(resp.response)['token'];

                    let authUrl = new URL(baseUrl + '/v1/auth/authorize');

                    let params = {
                        'token': token,
                        'clientId': clientId,
                        'serviceType': 'Office365',
                        'userAccount': 'primary',
                        'prompt': 'select_account',
                        'loginHint': hostMailbox.email,
                        'scopes': 'Mail.Read Mail.Send Calendar.ReadWrite Contacts.ReadWrite',
                    };

                    for (let paramsKey in params) {
                        authUrl.searchParams.append(paramsKey, params[paramsKey]);
                    }

                    let child = window.open(
                        authUrl.toString(),
                        'auth',
                        "width=800,height=600,resizable=1,scrollbars=1");

                    window.onmessage = function (event) {
                        console.log(event);
                        request('GET', baseUrl + '/v1/user', authHeaders, function (res) {
                            console.log(res);
                            let email = JSON.parse(res.response)['email'];
                            showUserEmail(email);
                            hideAuth();
                        });
                    };


                });


            }

        }
    });

    hostMailbox.email = Office.context.mailbox.userProfile.emailAddress;

};



function request(method, url, headers, ready) {
    var xhttp = new XMLHttpRequest();
    xhttp.onreadystatechange = function () {
        if (this.readyState === 4 && this.status === 200) {
            ready(this);
        }
    };

    xhttp.open(method, url, true);

    for (let headersKey in headers) {
        xhttp.setRequestHeader(headersKey, headers[headersKey]);
    }

    xhttp.send();
}

function showUserEmail(email) {
    var div = document.createElement('div');
    div.innerHTML = '<p>Welcome. You have successfully logged in.</p>'
    document.body.appendChild(div)
}

function hideAuth() {
    let auth = document.getElementById('auth');
    if (auth !== undefined) {
        auth.hidden = true;
    }
}