// Auth logic using SHA-256 for basic security

// Helper: SHA-256 hashing
async function hashPassword(password) {
    // Basic SHA-256 hash using Web Crypto API
    const msgBuffer = new TextEncoder().encode(password);
    const hashBuffer = await crypto.subtle.digest('SHA-256', msgBuffer);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
}

// Check if user is logged in
function checkAuth() {
    const session = sessionStorage.getItem('dashboard_session');
    if (!session) {
        // Not logged in
        return false;
    }
    
    try {
        const data = JSON.parse(session);
        // Simple expiry check
        if (!data.username || !data.expires || data.expires < Date.now()) {
            logout();
            return false;
        }
        return true;
    } catch (e) {
        logout();
        return false;
    }
}

// Login function
async function login(username, password) {
    try {
        // Load config
        const response = await fetch('assets/js/admin/config.json');
        if (!response.ok) {
            console.error('Failed to load config');
            return { success: false, message: '系统错误：无法加载配置文件' };
        }
        const config = await response.json();

        // Check username
        if (username !== config.username) {
            return { success: false, message: '用户名或密码错误' };
        }

        // Check password hash
        const inputHash = await hashPassword(password);
        if (inputHash !== config.password_hash) {
            return { success: false, message: '用户名或密码错误' };
        }

        // Success: Set session
        const sessionData = {
            username: username,
            loginTime: Date.now(),
            expires: Date.now() + (24 * 60 * 60 * 1000) // 24 hours
        };
        sessionStorage.setItem('dashboard_session', JSON.stringify(sessionData));
        return { success: true };

    } catch (e) {
        console.error('Login error:', e);
        return { success: false, message: '系统错误，请重试' };
    }
}

// Logout function
function logout() {
    sessionStorage.removeItem('dashboard_session');
    window.location.href = 'login.html';
}

// Export for module usage if needed, but here we use global scope for simple inclusion
window.Auth = {
    check: checkAuth,
    login: login,
    logout: logout
};
