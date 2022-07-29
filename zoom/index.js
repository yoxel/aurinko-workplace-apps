const clientIdName = 'clientId';
let baseUrl = window.location.origin + '/v1';
let params = new URLSearchParams(window.location.search);
let clientId = params.get(clientIdName);
let onAuthClick;
let onLogoutClick;
let getUser;


let authHeaders = {
    'X-Aurinko-ClientId': clientId,
    'content-type': 'application/json'
};

request('GET', baseUrl + '/user', authHeaders, function (res) {
    console.log(res);
    let parsed = JSON.parse(res.response);
    let email = parsed['email'];
    let name = parsed['accounts'][0]['name'];
    showUserInfo(email, name);
});

getUser = function () {
    request('GET', baseUrl + '/user', authHeaders, function (res) {
        console.log(res);
        let parsed = JSON.parse(res.response);
        let email = parsed['email'];
        let name = parsed['accounts'][0]['name'];
        showUserInfo(email, name);
    });
};

onAuthClick = function (isPopup) {
    let authUrl = new URL(baseUrl + '/auth/authorize');

    let params = {
        'clientId': clientId,
        'serviceType': 'Zoom',
        'userAccount': 'primary',
        'prompt': 'select_account',
        'loginHint': email,
    };

    if (!isPopup) {
        params['returnUrl'] = window.location.href;
    }

    for (let paramsKey in params) {
        authUrl.searchParams.append(paramsKey, params[paramsKey]);
    }

    if (isPopup) {
        let child = window.open(
            authUrl.toString(),
            '_blank',
            "width=800,height=600,resizable=1,scrollbars=1");

        window.onmessage = function (event) {
            console.log(event.data);
            request('GET', baseUrl + '/user', authHeaders, function (res) {
                console.log(res);
                let parsed = JSON.parse(res.response);
                let email = parsed['email'];
                let name = parsed['accounts'][0]['name'];
                showUserInfo(email, name);
            });
        };
    } else {
        window.location.href = authUrl.toString();
    }
};

function showUserInfo(email, name) {

    document.getElementById('authResult').innerHTML =
        '<label>User info</label>' +
        '<p class="">' + name + '</p>' +
        '<p>' + email + '</p>';
}

function request(method, url, headers, ready) {
    let xhttp = new XMLHttpRequest();
    xhttp.onreadystatechange = function () {
        if (this.readyState === 4 && this.status === 200) {
            ready(this);
        } else {
            console.log(this.responseText);
        }
    };

    xhttp.withCredentials = true;

    xhttp.open(method, url, true);

    for (let headersKey in headers) {
        xhttp.setRequestHeader(headersKey, headers[headersKey]);
    }

    xhttp.send();
}

