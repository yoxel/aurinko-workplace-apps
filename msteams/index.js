microsoftTeams.initialize()

let wlo = window.location.origin

let auClientId = getQueryParam('clientId')
var apiUrl = 'https://api.aurinko.io/v1'

let beforeAuthView, afterAuthView, errorView, loginInfo, usernameView, userEmailView


let auClientIdHeader = ['X-Aurinko-ClientId', auClientId]
let contentTypeHeader = ['Content-Type', 'application/json']

let auHeaders = [auClientIdHeader, contentTypeHeader]

document.addEventListener("DOMContentLoaded", function(e) {

    beforeAuthView = document.getElementById('unauthorized')
    afterAuthView = document.getElementById('authorized')
    errorView = document.getElementById('error-message')
    loginInfo = document.getElementById('login-info')
    usernameView = document.getElementById('username')
    userEmailView = document.getElementById('email')
})

window.onerror = function(e) {
    console.log(e)
    showError(e)
}


let authUrl = function() {
    return `${apiUrl}/auth/authorize/clientId=${auClientId}&serviceType=Office365&userAccount=primary&returnUrl=${wlo}/auth_callback.html`
}

function startAuthorization () {
    console.log('Auth start')

    microsoftTeams.authentication.authenticate({

        url: authUrl(),

        width: 600,
        height: 530,

        successCallback: function (result) {
            hideError()
            afterAuthView.style.display = 'block'
            showAccInfo()
        },
        
        failureCallback: function (reason) {
            showError(reason)
        }
    })
}

function showAccInfo() {
    const primaryAccount = auAccounts().find(acc => acc.primary)

    if(primaryAccount) {
        usernameView.innerHTML = primaryAccount.name
        userEmailView.innerHTML = primaryAccount.email

        loginInfo.style.display = block
    }
}

function auAccounts() {
    apiRequest('GET', '/user/accounts', auHeaders, function(response) {
        return response.records
    })
}

function apiRequest(method, path, headers, onSuccess) {
    var xhr = new XMLHttpRequest()

    xhr.open(method, `${apiUrl}${path}`)
    xhr.responseType = 'json'

    for (let header in headers) {
        xhr.setRequestHeader(...header)
    }

    xhr.onload = function() {
        if (200 <= xhr.status <= 300) {
            hideError()
            onSuccess(xhr.response)

        } else {
            throw `ApiException: ${xhr.status}. ${xhr.response.message}`
        }
    }

    xhr.send()
}


function showError(message) {
    errorView.innerHTML = message
    errorView.style.display = 'block'
}

function hideError() {
    errorView.innerHTML = ''
    errorView.style.display = 'none'
}


function getQueryParam (name) {
    const entry = window.location.search.substring(1).split('&').find(el => el.startsWith(name))

    if (!entry) {
        return null
    } else return entry.split('=')[1]
}
