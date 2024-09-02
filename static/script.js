// Obtener elementos
const modal = document.getElementById('modal');
const modalText = document.getElementById('modal-text');
const span = document.getElementsByClassName('close')[0];

// Mostrar el modal con el texto correspondiente
document.getElementById('soporteBtn').onclick = function() {
    modalText.innerHTML = "Para soporte o ayuda, puedes contactar al siguiente correo: jjamadorr@uaemex.mx";
    modal.style.display = "block";
}

document.getElementById('avisoBtn').onclick = function() {
    modalText.innerHTML = "Este es el aviso de privacidad:\n Tu información será tratada de acuerdo con nuestras políticas.";
    modal.style.display = "block";
}

document.getElementById('terminosBtn').onclick = function() {
    modalText.innerHTML = "Los términos de uso al usar este sitio:\n Tendras que aceptar nuestras condiciones para usar la aplicacion.";
    modal.style.display = "block";
}

// Cerrar cuando se hace clic en la 'X'
span.onclick = function() {
    modal.style.display = "none";
}

// Cerrar cuando se hace clic fuera del cuadro
window.onclick = function(event) {
    if (event.target == modal) {
        modal.style.display = "none";
    }
}
