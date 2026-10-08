document.addEventListener('DOMContentLoaded', () => {
    const mapAuthError = (code) => {
        switch (code) {
            case 'auth/user-not-found':
            case 'auth/invalid-credential':
            case 'auth/invalid-login-credentials':
            case 'auth/wrong-password':
                return 'Correo o contraseña incorrectos. Verifica que el usuario exista.';
            case 'auth/invalid-email':
                return 'El formato del correo electrónico no es válido.';
            case 'auth/too-many-requests':
                return 'Demasiados intentos fallidos. Por favor, intenta de nuevo más tarde.';
            case 'auth/email-already-in-use':
                return 'El correo electrónico ya está registrado en otra cuenta.';
            case 'auth/weak-password':
                return 'La contraseña es muy débil (mínimo 6 caracteres).';
            case 'auth/network-request-failed':
                return 'Error de conexión a internet.';
            default:
                return 'Ocurrió un error: ' + code;
        }
    };

    const loginOverlay = document.getElementById('loginOverlay');
    const loginForm = document.getElementById('loginForm');
    const loginEmail = document.getElementById('loginEmail');
    const loginPassword = document.getElementById('loginPassword');
    const loginError = document.getElementById('loginError');
    const logoutBtn = document.getElementById('logoutBtn');
    const appContainer = document.querySelector('.app-container');

    // Hide app container initially until auth is verified
    if (appContainer) {
        appContainer.style.display = 'none';
    }

    // Auth State Listener
    window.auth.onAuthStateChanged(async (user) => {
        if (user) {
            const emailDisplay = document.getElementById('currentUserEmailDisplay');
            if (emailDisplay) emailDisplay.textContent = user.email || 'Usuario';

            // MULTI-TENANT: Detect user change and clear old data
            window.currentUserTenant = user.uid;

            // CHECK ACCOUNT STATUS & REGISTER ROOT DOC
            try {
                if (typeof db !== 'undefined') {
                    const tenantDocRef = db.collection('tenants').doc(user.uid);
                    const tenantDoc = await tenantDocRef.get();
                    if (tenantDoc.exists) {
                        if (tenantDoc.data().status === 'suspended') {
                            await window.auth.signOut();
                            if (loginError) {
                                loginError.style.display = 'block';
                                loginError.textContent = 'Tu cuenta ha sido suspendida. Comunícate con soporte.';
                            }
                            return;
                        }
                        tenantDocRef.set({ lastLogin: firebase.firestore.FieldValue.serverTimestamp() }, { merge: true }).catch(e => console.error(e));
                        
                        // TRIAL LOGIC
                        let isExpired = false;
                        let remainingDays = 0;
                        const data = tenantDoc.data();
                        
                        let expirationDate = null;
                        if (data.trialEndsAt && typeof data.trialEndsAt.toDate === 'function') {
                            expirationDate = data.trialEndsAt.toDate();
                        } else if (data.createdAt && typeof data.createdAt.toDate === 'function') {
                            expirationDate = new Date(data.createdAt.toDate().getTime() + (15 * 24 * 60 * 60 * 1000));
                        }

                        if (expirationDate) {
                            const now = new Date();
                            const diffMs = expirationDate.getTime() - now.getTime();
                            remainingDays = Math.ceil(diffMs / (1000 * 60 * 60 * 24));
                            if (remainingDays < 0) remainingDays = 0;
                            
                            if (now > expirationDate) {
                                isExpired = true;
                            }
                        } else {
                            // Si es una cuenta tan nueva que an no tiene createdAt procesado en el servidor
                            remainingDays = 15;
                        }

                        if (isExpired) {
                            const trialOverlay = document.getElementById('trialExpiredOverlay');
                            if (trialOverlay) trialOverlay.style.display = 'flex';
                            if (loginOverlay) loginOverlay.style.display = 'none';
                            if (appContainer) appContainer.style.display = 'none';
                            return; // Stop loading app
                        } else {
                            const counter = document.getElementById('trialRemainingCounter');
                            if (counter) counter.innerHTML = "Quedan <b>" + remainingDays + " d&iacute;as</b> de prueba";
                        }
                    } else {
                        // First login: register tenant metadata for the Super Admin panel
                        await tenantDocRef.set({
                            email: user.email,
                            createdAt: firebase.firestore.FieldValue.serverTimestamp(),
                            lastLogin: firebase.firestore.FieldValue.serverTimestamp(),
                            status: 'active'
                        }, { merge: true });
                        const counter = document.getElementById('trialRemainingCounter');
                        if (counter) counter.innerHTML = "Quedan <b>15 d&iacute;as</b> de prueba";
                    }
                }
            } catch(e) {
                console.error("Error validando cuenta:", e);
            }

            const lastTenant = localStorage.getItem('minesof_last_tenant');
            const isNewUserLocal = (!lastTenant || lastTenant !== user.uid);
            
            if (isNewUserLocal) {
                StorageManager.clearAll();
                localStorage.removeItem('galeria_admin_password');
                localStorage.removeItem('galeria_observations');
                localStorage.setItem('minesof_last_tenant', user.uid);
                
                const pendingBilling = localStorage.getItem('minesof_pending_registration_billing');
                if (pendingBilling && typeof db !== 'undefined') {
                    localStorage.removeItem('minesof_pending_registration_billing');
                    db.collection('tenants').doc(user.uid).collection('minesof_settings').doc('global_config').set({
                        billingSystem: pendingBilling
                    }, { merge: true });
                    localStorage.setItem('minesof_billingSystem', pendingBilling);
                }
            }

            // Reload the configuration for this specific tenant
            const config = StorageManager.getConfig();
            Object.assign(FOODX_DATA, config);

            // Mostrar la app inmediatamente, sin pantalla de "Sincronizando..."
            loginOverlay.style.display = 'none';
            if (appContainer) appContainer.style.display = 'flex';
        } else {
            // User is signed out
            loginOverlay.style.display = 'flex';
            if (appContainer) appContainer.style.display = 'none';
        }
    });

    const forgotPasswordLink = document.getElementById('forgotPasswordLink');
    const confirmPasswordGroup = document.getElementById('confirmPasswordGroup');
    const loginConfirmPassword = document.getElementById('loginConfirmPassword');

    if (forgotPasswordLink) {
        forgotPasswordLink.addEventListener('click', (e) => {
            e.preventDefault();
            const email = loginEmail.value.trim();
            const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
            
            if (!email || !emailRegex.test(email)) {
                alert('Por favor, ingresa un correo electrónico válido en el campo de arriba para enviarte el enlace de recuperación (ejemplo: usuario@correo.com).');
                return;
            }
            window.auth.languageCode = 'es';
            window.auth.sendPasswordResetEmail(email)
                .then(() => {
                    alert('¡Correo enviado a ' + email + '!\n\n1. Revisa tu bandeja (o Spam) y haz clic en el enlace para crear tu nueva contraseña.\n2. Luego, regresa a esta pantalla e inicia sesión normalmente.');
                    document.getElementById('loginPassword').value = '';
                })
                .catch((error) => {
                    alert('Error: ' + mapAuthError(error.code));
                });
        });
    }

    const authToggleLink = document.getElementById('authToggleLink');
    const btn = document.getElementById('loginBtn');
    let isLoginMode = true;

    if (authToggleLink) {
        authToggleLink.addEventListener('click', (e) => {
            e.preventDefault();
            isLoginMode = !isLoginMode;
            if (isLoginMode) {
                authTitle.textContent = 'Bienvenido a Minesof';
                if (authSubtitle) authSubtitle.textContent = 'Inicia sesi\u00F3n para acceder a tu sistema';
                btn.textContent = 'Ingresar';
                authToggleLink.innerHTML = '&iquest;No tienes cuenta? Reg&iacute;strate aqu&iacute;';
                if (forgotPasswordLink && forgotPasswordLink.parentElement) forgotPasswordLink.parentElement.style.display = 'block';
                if (confirmPasswordGroup) confirmPasswordGroup.style.display = 'none';
                const confirmEmailGroup = document.getElementById('confirmEmailGroup');
                if (confirmEmailGroup) confirmEmailGroup.style.display = 'none';
                const loginConfirmEmail = document.getElementById('loginConfirmEmail');
                if (loginConfirmEmail) loginConfirmEmail.value = '';
                const registerBillingGroup = document.getElementById('registerBillingGroup');
                if (registerBillingGroup) registerBillingGroup.style.display = 'none';
                if (loginConfirmPassword) loginConfirmPassword.value = '';
            } else {
                authTitle.textContent = 'Crear Nueva Cuenta';
                if (authSubtitle) authSubtitle.textContent = 'Crea una cuenta para empezar a usar el sistema';
                btn.textContent = 'Registrarse';
                authToggleLink.innerHTML = '&iquest;Ya tienes cuenta? Inicia Sesi&oacute;n';
                if (forgotPasswordLink && forgotPasswordLink.parentElement) forgotPasswordLink.parentElement.style.display = 'none';
                if (confirmPasswordGroup) confirmPasswordGroup.style.display = 'block';
                const confirmEmailGroup = document.getElementById('confirmEmailGroup');
                if (confirmEmailGroup) confirmEmailGroup.style.display = 'block';
                const registerBillingGroup = document.getElementById('registerBillingGroup');
                if (registerBillingGroup) registerBillingGroup.style.display = 'block';
            }
        });
    }

    // Toggle Password Visibility
    document.querySelectorAll('.toggle-password').forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.preventDefault();
            const targetId = btn.getAttribute('data-target');
            const input = document.getElementById(targetId);

            if (input) {
                if (input.type === 'password') {
                    input.type = 'text';
                    btn.innerHTML = '<i data-lucide="eye-off" style="width: 20px; height: 20px; color: #000000; font-weight: 600;"></i>';
                } else {
                    input.type = 'password';
                    btn.innerHTML = '<i data-lucide="eye" style="width: 20px; height: 20px; color: #000000; font-weight: 600;"></i>';
                }
                if (window.lucide) {
                    window.lucide.createIcons();
                }
            }
        });
    });

    // Login Submit
    loginForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const email = loginEmail.value.trim();
        const password = loginPassword.value;

        loginError.style.display = 'none';
        
        // Validacion estricta de correo (debe tener algo@algo.com)
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailRegex.test(email)) {
            loginError.textContent = 'Por favor, ingresa un correo electrónico válido (ejemplo: usuario@correo.com).';
            loginError.style.display = 'block';
            return;
        }

        btn.textContent = 'Procesando...';
        btn.disabled = true;

        if (!isLoginMode) {
            const domain = email.split('@')[1];
            const allowedDomains = ['gmail.com', 'hotmail.com', 'outlook.com', 'yahoo.com', 'yahoo.es', 'hotmail.es', 'outlook.es'];
            
            if (!domain || !allowedDomains.includes(domain.toLowerCase())) {
                loginError.textContent = 'Por seguridad, solo aceptamos correos oficiales (@gmail.com, @hotmail.com, @outlook.com o @yahoo.com). Verifica que esté bien escrito.';
                loginError.style.display = 'block';
                btn.textContent = 'Registrarse';
                btn.disabled = false;
                return;
            }

            const confirmEmailInput = document.getElementById('loginConfirmEmail');
            if (confirmEmailInput && confirmEmailInput.value.trim() !== email) {
                loginError.textContent = 'Los correos electrónicos no coinciden.';
                loginError.style.display = 'block';
                btn.textContent = 'Registrarse';
                btn.disabled = false;
                return;
            }

            const confirmPass = loginConfirmPassword ? loginConfirmPassword.value : '';
            if (password !== confirmPass) {
                loginError.textContent = 'Las contraseñas no coinciden.';
                loginError.style.display = 'block';
                btn.textContent = 'Registrarse';
                btn.disabled = false;
                return;
            }
        }

        if (isLoginMode) {
            window.auth.signInWithEmailAndPassword(email, password)
                .then((userCredential) => {
                    const config = StorageManager.getConfig();
                    config.adminPassword = password;
                    StorageManager.saveConfig(config);
                    // La pantalla se oculta automticamente por onAuthStateChanged
                    // No limpiamos el formulario para evitar parpadeos visuales
                })
                .catch((error) => {
                    loginError.textContent = mapAuthError(error.code);
                    loginError.style.display = 'block';
                    btn.textContent = 'Ingresar';
                    btn.disabled = false;
                });
        } else {
            const regBillingSystem = document.getElementById('registerBillingSystem');
            if (regBillingSystem) {
                localStorage.setItem('minesof_pending_registration_billing', regBillingSystem.value);
            }
            window.auth.createUserWithEmailAndPassword(email, password)
                .then((userCredential) => {
                    const config = StorageManager.getConfig();
                    config.adminPassword = password;
                    StorageManager.saveConfig(config);
                    // La pantalla se oculta automticamente por onAuthStateChanged
                    // No limpiamos el formulario para evitar parpadeos visuales
                })
                .catch((error) => {
                    localStorage.removeItem('minesof_pending_registration_billing');
                    loginError.textContent = mapAuthError(error.code);
                    loginError.style.display = 'block';
                    btn.textContent = 'Registrarse';
                    btn.disabled = false;
                });
        }
    });

    // Logout Click
    if (logoutBtn) {
        logoutBtn.addEventListener('click', () => {
            StorageManager.clearAll();
            localStorage.removeItem('minesof_last_tenant');
            window.auth.signOut().then(() => {
                window.location.reload();
            });
        });
    }
});
