microsoftTeams.initialize()

window.onerror = function(e) {
    console.log(e)
    showError(e)
}

let wlo = window.location.origin

let auClientId = getQueryParam('clientId')
// var apiUrl = 'https://api.aurinko.io/v1'
var apiUrl = wlo + '/v1'

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


let authUrl = function() {
    return `${wlo}/v1/auth/authorize?clientId=${auClientId}&serviceType=Office365&userAccount=primary&returnUrl=${encodeURIComponent(wlo + '/msteams/auth_callback.html')}`
}

function startAuthorization () {
    console.log('Auth start')

    microsoftTeams.authentication.authenticate({

        url: authUrl(),

        width: 600,
        height: 530,

        successCallback: function (result) {
            console.log('Auth success')
            hideError()
            beforeAuthView.style.display = 'none'
            afterAuthView.style.display = 'block'

            showAccInfo()
        },
        
        failureCallback: function (reason) {
            console.log('Auth failure')
            showError(reason)
        }
    })
}

function showAccInfo() {
    auAccounts(function(accounts){
        const primaryAccount = accounts.find(acc => acc.userAccountType == 'primary')

        if(primaryAccount) {
            usernameView.innerHTML = primaryAccount.name
            userEmailView.innerHTML = primaryAccount.email
    
            loginInfo.style.display = 'block'
        }
    })
}

function auAccounts(callback) {
    
    apiRequest('GET', '/user/accounts', auHeaders, function(response) {
        callback(response.records)
    })
}

function apiRequest(method, path, headers, onSuccess) {
    var xhr = new XMLHttpRequest()

    xhr.open(method, `${apiUrl}${path}`)
    xhr.responseType = 'json'


    for (let header of headers) {
        xhr.setRequestHeader(...header)
    }

    xhr.onload = function() {
        if (xhr.status >= 200 && xhr.status <= 300) {
            hideError()
            onSuccess(xhr.response)

        } else {
            throw `ApiException: ${xhr.status}. ${xhr.response.message}`
        }
    }

    xhr.send()
}


function showError(message) {
    console.log('showError ' + message)
    if (message){
        errorView.innerHTML = message
    } else {
        errorView.innerHTML = 'Unknown error.'
    }
    
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

