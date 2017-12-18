$("#loginBtn").click(function () {
    var selected = localStorage.getItem("selectedLanguage")
    window.location.href = "Account/SignIn?Lang=" + selected
})
