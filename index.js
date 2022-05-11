// microsoftTeams.initialize()

let wlo = window.location.origin

let auClientId = getQueryParam('clientId')

let beforeAuthView
let afterAuthView
let errorView

document.addEventListener("DOMContentLoaded", function(event) {

    beforeAuthView = document.getElementById('unauthorized')
    afterAuthView = document.getElementById('authorized')
    errorView = document.getElementById('error-message')
  })

let authUrl = function() {
    return `https://api.aurinko.io/auth/authorize/clientId=${auClientId}&serviceType=Office365&userAccount=primary&returnUrl=${wlo}/auth_callback.html`
}

function startAuthorization () {
    console.log('Auth start')

    microsoftTeams.authentication.authenticate({

        url: authUrl(),

        width: 600,
        height: 530,

        successCallback: function (result) {
            beforeAuthView.style.display = 'none'
            errorView.style.display = 'none'

            afterAuthView.style.display = 'block'
        },
        
        failureCallback: function (reason) {

            errorView.innerHTML = reason
            errorView.style.display = 'block'

        }
    })
}

window.onerror = function(e) {
    console.log(e)
    errorView.innerHTML = e
    errorView.style.display = 'block'
}

function getQueryParam (name, url = window.location.href) {
    name = name.replace(/[\[\]]/g, '\\$&');

    var regex = new RegExp('[?&]' + name + '(=([^&#]*)|&|#|$)')
    var results = regex.exec(url)

    if (!results) return null;
    if (!results[2]) return '';
    return decodeURIComponent(results[2].replace(/\+/g, ' '));
}
