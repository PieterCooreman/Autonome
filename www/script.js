// ====================== Autonome Agentic CMS - React SPA ======================
const { useState, useEffect, useCallback, useRef } = React;

// ====================== API HELPERS ======================
const api = async (action, method = 'GET', body = null) => {
  const opts = { method, headers: {} };
  if (method === 'POST') {
    if (body instanceof FormData) {
      opts.body = body;
    } else if (body) {
      opts.headers['Content-Type'] = 'application/x-www-form-urlencoded';
      opts.body = 'json=' + encodeURIComponent(JSON.stringify(body));
    }
  }
  const res = await fetch(`api.asp?action=${action}`, opts);
  const data = await res.json();
  if (!data.success) throw new Error(data.error || 'Unknown error');
  return data.data;
};

const apiRaw = async (action, method = 'GET', body = null) => {
  const opts = { method, headers: {} };
  if (method === 'POST' && body) {
    opts.headers['Content-Type'] = 'application/x-www-form-urlencoded';
    opts.body = 'json=' + encodeURIComponent(JSON.stringify(body));
  }
  const res = await fetch(`api.asp?action=${action}`, opts);
  return res.json();
};

const sleep = (ms) => new Promise(r => setTimeout(r, ms));

const h = React.createElement;

// ====================== APP ======================
function App() {
  const [user, setUser] = useState(null);
  const [loading, setLoading] = useState(true);
  const [route, setRoute] = useState(window.location.hash.slice(1) || '/');
  const [toast, setToast] = useState(null);

  const showToast = useCallback((message, type = 'success') => {
    setToast({ message, type });
    setTimeout(() => setToast(null), 2500);
  }, []);

  useEffect(() => {
    const onHashChange = () => setRoute(window.location.hash.slice(1) || '/');
    window.addEventListener('hashchange', onHashChange);
    api('session').then(u => { if (u) setUser(u); }).catch(() => {}).finally(() => setLoading(false));
    return () => window.removeEventListener('hashchange', onHashChange);
  }, []);

  // Keep session alive: ping every 5 minutes while logged in so sessions do not time out
  useEffect(() => {
    if (!user) return;
    const iv = setInterval(() => {
      api('keepalive').catch(() => {});
    }, 5 * 60 * 1000);
    return () => clearInterval(iv);
  }, [user]);

  const handleLogin = (userData) => setUser(userData);
  const handleLogout = async () => {
    await api('logout');
    setUser(null);
    window.location.hash = '#/';
  };

  const uploadImageRaw = async (files, projectId) => {
    const fd = new FormData();
    fd.append('project_id', projectId);
    const fileList = Array.isArray(files) ? files : [files];
    fileList.forEach(f => fd.append('files', f));
    const res = await fetch('api.asp?action=upload_image', { method: 'POST', body: fd });
    const data = await res.json();
    if (!data.success) throw new Error(data.error);
    return data.data;
  };

  if (loading) {
    return h('div', { className: 'min-h-screen flex items-center justify-center' },
      h('div', { className: 'text-center' },
        h('div', { className: 'spinner mx-auto mb-4', style: { width: 40, height: 40 } }),
        h('p', { style: { color: 'var(--text-dim)' } }, 'Loading...')
      )
    );
  }

  const renderRoute = () => {
    if (route.startsWith('/reset/')) return h(ResetPasswordPage, { token: route.split('/')[2] || '', showToast });
    if (user && route === '/admin') return h(AdminPanel, { user, showToast, onLogout: handleLogout });
    if (user && route === '/account') return h(AccountPage, { user, showToast, onLogout: handleLogout, onUserUpdate: setUser });
    if (user && route.startsWith('/project/')) return h(ProjectWorkspace, { user, projectId: route.split('/')[2], showToast, uploadImageRaw });
    if (user && route !== '/' && route !== '') return h(UserDashboard, { user, showToast, onLogout: handleLogout });
    if (user) return h(UserDashboard, { user, showToast, onLogout: handleLogout });
    if (route === '/login') return h(LoginPage, { onLogin: handleLogin, showToast });
    return h(LandingPage);
  };

  return h('div', { className: 'min-h-screen' },
    renderRoute(),
    toast && h('div', { className: `toast toast-${toast.type}` }, toast.message)
  );
}

// ====================== LANDING PAGE ======================
function LandingPage() {
  const goLogin = () => window.location.hash = '#/login';

  const feature = (icon, title, text) => h('div', { className: 'card', style: { padding: 22 } },
    h('div', { style: { fontSize: 30, marginBottom: 10 } }, icon),
    h('h3', { className: 'font-semibold text-white mb-1', style: { fontSize: 16 } }, title),
    h('p', { className: 'text-sm', style: { color: 'var(--text-dim)', lineHeight: 1.55 } }, text)
  );

  const step = (nr, title, text) => h('div', { className: 'flex gap-4 items-start' },
    h('div', { style: { width: 40, height: 40, borderRadius: 12, background: 'linear-gradient(135deg, var(--accent), #ea580c)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 800, fontSize: 17, color: '#fff', flexShrink: 0, boxShadow: '0 4px 14px rgba(249,115,22,0.35)' } }, nr),
    h('div', null,
      h('h4', { className: 'font-semibold text-white mb-1' }, title),
      h('p', { className: 'text-sm', style: { color: 'var(--text-dim)' } }, text)
    )
  );

  return h('div', { className: 'min-h-screen' },
    // Nav
    h('header', { className: 'gradient-header px-6 py-3' },
      h('div', { className: 'max-w-6xl mx-auto flex items-center justify-between' },
        h('div', { className: 'flex items-center gap-3' },
          h('div', { className: 'app-logo' }, '\u{1F916}'),
          h('span', { className: 'font-bold text-white', style: { fontSize: 17 } }, 'Autonome ', h('span', { style: { color: 'var(--accent)' } }, 'Agentic CMS'))
        ),
        h('button', { onClick: goLogin, className: 'gradient-btn px-5 py-2 text-white font-semibold rounded-xl text-sm' }, 'Sign in')
      )
    ),

    // Hero
    h('section', { className: 'max-w-4xl mx-auto text-center px-6', style: { paddingTop: 90, paddingBottom: 70 } },
      h('div', { className: 'animate-slide-up' },
        h('span', { className: 'badge badge-admin', style: { marginBottom: 22, display: 'inline-block' } }, 'AI-powered \u00B7 self-hosted \u00B7 private'),
        h('h1', { className: 'font-bold text-white', style: { fontSize: 52, lineHeight: 1.1, marginBottom: 22 } },
          'Websites that ', h('span', { style: { background: 'linear-gradient(135deg, var(--accent), #fbbf24)', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent' } }, 'build themselves'), '.'
        ),
        h('p', { className: 'text-lg mx-auto', style: { color: 'var(--text-dim)', maxWidth: 620, marginBottom: 36 } },
          'Autonome Agentic CMS is a content management system without the management. Describe your website in plain language \u2014 an autonomous AI agent researches, designs and writes the complete site. On your own hardware, with your own AI model.'
        ),
        h('div', { className: 'flex gap-3 justify-center' },
          h('button', { onClick: goLogin, className: 'gradient-btn px-8 py-3.5 text-white font-bold rounded-xl' }, 'Start building \u2192'),
          h('a', { href: '#how', className: 'btn-ghost px-8 py-3.5 rounded-xl font-semibold', style: { textDecoration: 'none' } }, 'How it works')
        )
      )
    ),

    // Features
    h('section', { className: 'max-w-6xl mx-auto px-6', style: { paddingBottom: 80 } },
      h('div', { className: 'grid grid-cols-1 md:grid-cols-3 gap-5' },
        feature('\u{1F9E0}', 'Agentic AI generation', 'Not just text completion: the agent researches your topic (Wikipedia included) and writes the complete site for you.'),
        feature('\u{1F512}', 'Your hardware, your data', 'Runs against your local LLM (LM Studio, Ollama). No cloud, no subscriptions, no data leaving your network.'),
        feature('\u{1F4F7}', 'Bring your photos', 'Upload images and the agent weaves them into the design. Automatic resizing included.'),
        feature('\u23F1\uFE0F', 'Version history', 'Every generation is backed up automatically. Preview and restore any previous version with one click.'),
        feature('\u{1F3A8}', 'Bootstrap 5 output', 'Clean, responsive, semantic HTML/CSS/JS built on Bootstrap 5.3 \u2014 mobile-first, with cookie & privacy notices included. Downloadable as a zip, ready to publish anywhere.'),
        feature('\u{1F465}', 'Multi-user', 'Every user gets isolated projects and files. Admins manage the AI configuration centrally.')
      )
    ),

    // How it works
    h('section', { id: 'how', className: 'max-w-3xl mx-auto px-6', style: { paddingBottom: 90 } },
      h('h2', { className: 'text-2xl font-bold text-white text-center mb-10' }, 'How it works'),
      h('div', { className: 'flex flex-col gap-8' },
        step('1', 'Create a project', 'Sign in, create a project \u2014 your site gets its own URL instantly.'),
        step('2', 'Upload photos & describe your site', 'Add your images and tell the agent what you want: purpose, audience, style, content.'),
        step('3', 'The agent builds it', 'The AI researches your topic, writes the code, and publishes the site. Keep refining in plain language \u2014 move sections, add content, restyle \u2014 or roll back any time.')
      )
    ),

    // CTA + footer
    h('section', { className: 'max-w-3xl mx-auto text-center px-6', style: { paddingBottom: 60 } },
      h('div', { className: 'card', style: { padding: 40 } },
        h('h2', { className: 'text-2xl font-bold text-white mb-2' }, 'Ready to build without building?'),
        h('p', { className: 'mb-6', style: { color: 'var(--text-dim)' } }, 'Create a free account and generate your first website in minutes.'),
        h('button', { onClick: goLogin, className: 'gradient-btn px-8 py-3.5 text-white font-bold rounded-xl' }, 'Get started')
      )
    ),
    h('footer', { className: 'text-center py-8 text-xs', style: { color: 'var(--text-mute)', borderTop: '1px solid var(--border)' } },
      'Autonome Agentic CMS \u2014 autonomous website generation on your own terms.'
    )
  );
}

// ====================== LOGIN / REGISTER / FORGOT ======================
function LoginPage({ onLogin }) {
  const [mode, setMode] = useState('login'); // 'login' | 'register' | 'forgot'
  const [username, setUsername] = useState('');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [info, setInfo] = useState('');

  const switchMode = (m) => { setMode(m); setError(''); setInfo(''); };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError('');
    setInfo('');

    if (mode === 'forgot') {
      if (!email.trim()) { setError('Email address required'); return; }
      setLoading(true);
      try {
        const res = await apiRaw('forgot_password', 'POST', { email: email.trim() });
        if (res.success) {
          setInfo(typeof res.data === 'string' ? res.data : 'If that email address is registered, a reset link has been sent.');
        } else {
          setError(res.error);
        }
      } catch (err) { setError(err.message); }
      finally { setLoading(false); }
      return;
    }

    if (!username.trim() || !password.trim()) { setError('All fields required'); return; }
    if (mode === 'register') {
      if (!email.trim()) { setError('Email address required'); return; }
      if (password.length < 6) { setError('Password must be at least 6 characters'); return; }
    }
    setLoading(true);
    try {
      const action = mode === 'register' ? 'register' : 'login';
      const payload = { username: username.trim(), password };
      if (mode === 'register') payload.email = email.trim();
      const res = await apiRaw(action, 'POST', payload);
      if (res.success) {
        onLogin(res.data);
        window.location.hash = res.data.role === 'admin' && mode === 'login' ? '#/admin' : '#/dashboard';
      } else {
        setError(res.error);
      }
    } catch (err) { setError(err.message); }
    finally { setLoading(false); }
  };

  const fieldBox = (label, input) => h('div', { className: 'mb-4' },
    h('label', { className: 'section-label block mb-2' }, label),
    input
  );

  return h('div', { className: 'min-h-screen flex items-center justify-center p-4' },
    h('div', { className: 'w-full animate-slide-up', style: { maxWidth: 420 } },
      h('div', { className: 'text-center mb-8' },
        h('div', { className: 'app-logo mx-auto mb-4', style: { width: 60, height: 60, fontSize: 30, borderRadius: 18, cursor: 'pointer' }, onClick: () => window.location.hash = '#/' }, '\u{1F916}'),
        h('h1', { className: 'text-3xl font-bold text-white mb-1' }, 'Autonome ', h('span', { style: { color: 'var(--accent)' } }, 'Agentic CMS')),
        h('p', { style: { color: 'var(--text-dim)' }, className: 'text-sm' }, 'Describe it. The agent builds it.')
      ),
      h('div', { className: 'card' },
        h('h2', { className: 'text-lg font-semibold text-white mb-5' },
          mode === 'register' ? 'Create your account' : mode === 'forgot' ? 'Forgot your password?' : 'Welcome back'),
        error && h('div', { className: 'px-4 py-3 rounded-lg mb-4 text-sm', style: { background: 'rgba(239,68,68,0.1)', border: '1px solid rgba(239,68,68,0.35)', color: '#fca5a5' } }, error),
        info && h('div', { className: 'px-4 py-3 rounded-lg mb-4 text-sm', style: { background: 'rgba(16,185,129,0.1)', border: '1px solid rgba(16,185,129,0.35)', color: '#6ee7b7' } }, info),
        mode === 'forgot' && h('p', { className: 'text-sm mb-4', style: { color: 'var(--text-dim)' } },
          'Enter the email address of your account and we will send you a link to set a new password.'),
        h('form', { onSubmit: handleSubmit },
          mode !== 'forgot' && fieldBox('Username',
            h('input', { type: 'text', value: username, onChange: e => setUsername(e.target.value), placeholder: 'your-username', className: 'w-full', autoComplete: 'username' })),
          (mode === 'register' || mode === 'forgot') && fieldBox('Email address',
            h('input', { type: 'email', value: email, onChange: e => setEmail(e.target.value), placeholder: 'you@example.com', className: 'w-full', autoComplete: 'email' })),
          mode !== 'forgot' && h('div', { className: 'mb-2' },
            h('label', { className: 'section-label block mb-2' }, 'Password'),
            h('input', { type: 'password', value: password, onChange: e => setPassword(e.target.value), placeholder: '\u2022\u2022\u2022\u2022\u2022\u2022\u2022\u2022', className: 'w-full', autoComplete: mode === 'register' ? 'new-password' : 'current-password' })
          ),
          mode === 'login' && h('div', { className: 'mb-4 text-right' },
            h('button', { type: 'button', onClick: () => switchMode('forgot'), className: 'text-xs', style: { color: 'var(--text-dim)', background: 'none', border: 'none', cursor: 'pointer' } }, 'Forgot password?')
          ),
          h('button', { type: 'submit', disabled: loading, className: 'gradient-btn w-full py-3 px-4 text-white font-semibold rounded-xl disabled:opacity-50 mt-2' },
            loading ? 'Please wait...' : (mode === 'register' ? 'Create Account' : mode === 'forgot' ? 'Send reset link' : 'Sign In')
          )
        ),
        h('div', { className: 'mt-5 text-center' },
          mode === 'forgot'
            ? h('button', { onClick: () => switchMode('login'), className: 'text-sm transition-colors', style: { color: 'var(--accent-2)', background: 'none', border: 'none', cursor: 'pointer' } }, '\u2190 Back to sign in')
            : h('button', { onClick: () => switchMode(mode === 'register' ? 'login' : 'register'), className: 'text-sm transition-colors', style: { color: 'var(--accent-2)', background: 'none', border: 'none', cursor: 'pointer' } },
                mode === 'register' ? 'Already have an account? Sign in' : "No account yet? Register free"
              )
        )
      )
    )
  );
}

// ====================== RESET PASSWORD (from email link) ======================
function ResetPasswordPage({ token, showToast }) {
  const [password, setPassword] = useState('');
  const [confirm, setConfirm] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [done, setDone] = useState(false);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError('');
    if (password.length < 6) { setError('Password must be at least 6 characters'); return; }
    if (password !== confirm) { setError('Passwords do not match'); return; }
    setLoading(true);
    try {
      const res = await apiRaw('reset_password', 'POST', { token, password });
      if (res.success) {
        setDone(true);
        showToast('Password reset. You can now sign in.');
      } else {
        setError(res.error);
      }
    } catch (err) { setError(err.message); }
    finally { setLoading(false); }
  };

  return h('div', { className: 'min-h-screen flex items-center justify-center p-4' },
    h('div', { className: 'w-full animate-slide-up', style: { maxWidth: 420 } },
      h('div', { className: 'text-center mb-8' },
        h('div', { className: 'app-logo mx-auto mb-4', style: { width: 60, height: 60, fontSize: 30, borderRadius: 18 } }, '\u{1F511}'),
        h('h1', { className: 'text-2xl font-bold text-white mb-1' }, 'Set a new password')
      ),
      h('div', { className: 'card' },
        done
          ? h('div', { className: 'text-center' },
              h('p', { className: 'mb-5', style: { color: 'var(--text-dim)' } }, 'Your password has been changed successfully.'),
              h('button', { onClick: () => window.location.hash = '#/login', className: 'gradient-btn px-6 py-3 text-white font-semibold rounded-xl' }, 'Go to sign in')
            )
          : h('form', { onSubmit: handleSubmit },
              error && h('div', { className: 'px-4 py-3 rounded-lg mb-4 text-sm', style: { background: 'rgba(239,68,68,0.1)', border: '1px solid rgba(239,68,68,0.35)', color: '#fca5a5' } }, error),
              h('div', { className: 'mb-4' },
                h('label', { className: 'section-label block mb-2' }, 'New password'),
                h('input', { type: 'password', value: password, onChange: e => setPassword(e.target.value), className: 'w-full', autoComplete: 'new-password' })
              ),
              h('div', { className: 'mb-6' },
                h('label', { className: 'section-label block mb-2' }, 'Confirm new password'),
                h('input', { type: 'password', value: confirm, onChange: e => setConfirm(e.target.value), className: 'w-full', autoComplete: 'new-password' })
              ),
              h('button', { type: 'submit', disabled: loading, className: 'gradient-btn w-full py-3 px-4 text-white font-semibold rounded-xl disabled:opacity-50' },
                loading ? 'Please wait...' : 'Set new password')
            )
      )
    )
  );
}

// ====================== HEADER ======================
function AppHeader({ user, onLogout, subtitle, extra }) {
  return h('header', { className: 'gradient-header px-6 py-3' },
    h('div', { className: 'max-w-6xl mx-auto flex items-center justify-between' },
      h('div', { className: 'flex items-center gap-3 cursor-pointer', onClick: () => window.location.hash = '#/dashboard' },
        h('div', { className: 'app-logo' }, '\u{1F916}'),
        h('div', null,
          h('h1', { className: 'font-bold text-white header-title', style: { fontSize: 16, lineHeight: 1.2 } }, 'Autonome Agentic CMS'),
          subtitle && h('p', { className: 'text-xs header-subtitle', style: { color: 'var(--text-dim)' } }, subtitle)
        )
      ),
      h('div', { className: 'flex items-center gap-3 nav-actions' },
        extra,
        user.role === 'admin' && h('button', { onClick: () => window.location.hash = '#/admin', className: 'btn-ghost px-3 py-2 rounded-lg text-sm nav-btn' }, h('span', { className: 'nav-icon' }, '\u2699\uFE0F'), h('span', { className: 'nav-label' }, ' Admin')),
        h('button', { onClick: () => window.location.hash = '#/account', className: 'btn-ghost px-3 py-2 rounded-lg text-sm nav-btn' }, h('span', { className: 'nav-icon' }, '\u{1F464}'), h('span', { className: 'nav-label' }, ' Account')),
        h('button', { onClick: () => window.location.hash = '#/dashboard', className: 'btn-ghost px-3 py-2 rounded-lg text-sm nav-btn' }, h('span', { className: 'nav-icon' }, '\u{1F3E0}'), h('span', { className: 'nav-label' }, ' Dashboard')),
        h('button', { onClick: onLogout, className: 'btn-ghost px-3 py-2 rounded-lg text-sm nav-btn' }, h('span', { className: 'nav-icon' }, '\u238B'), h('span', { className: 'nav-label' }, ' Logout'))
      )
    )
  );
}

// ====================== ACCOUNT SETTINGS ======================
function AccountPage({ user, showToast, onLogout, onUserUpdate }) {
  const [emailPassword, setEmailPassword] = useState('');
  const [newEmail, setNewEmail] = useState(user.email || '');
  const [emailSaving, setEmailSaving] = useState(false);

  const [curPassword, setCurPassword] = useState('');
  const [newPassword, setNewPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [pwSaving, setPwSaving] = useState(false);

  const handleEmailSave = async (e) => {
    e.preventDefault();
    if (!emailPassword || !newEmail.trim()) { showToast('Password and new email required', 'error'); return; }
    setEmailSaving(true);
    try {
      const res = await api('update_email', 'POST', { password: emailPassword, email: newEmail.trim() });
      showToast('Email address updated');
      setEmailPassword('');
      onUserUpdate({ ...user, email: res.email });
    } catch (err) { showToast(err.message, 'error'); }
    finally { setEmailSaving(false); }
  };

  const handlePasswordSave = async (e) => {
    e.preventDefault();
    if (!curPassword || !newPassword) { showToast('All password fields required', 'error'); return; }
    if (newPassword.length < 6) { showToast('New password must be at least 6 characters', 'error'); return; }
    if (newPassword !== confirmPassword) { showToast('New passwords do not match', 'error'); return; }
    setPwSaving(true);
    try {
      await api('change_password', 'POST', { current_password: curPassword, new_password: newPassword });
      showToast('Password changed');
      setCurPassword(''); setNewPassword(''); setConfirmPassword('');
    } catch (err) { showToast(err.message, 'error'); }
    finally { setPwSaving(false); }
  };

  const field = (label, input) => h('div', { className: 'mb-5' },
    h('label', { className: 'section-label block mb-2' }, label),
    input
  );

  return h('div', { className: 'min-h-screen' },
    h(AppHeader, { user, onLogout, subtitle: 'Account settings' }),
    h('main', { className: 'max-w-2xl mx-auto p-6' },
      h('h2', { className: 'text-2xl font-bold text-white mb-6 mt-2' }, 'Account settings'),

      h('div', { className: 'card mb-6' },
        h('h3', { className: 'text-lg font-semibold text-white mb-1' }, 'Email address'),
        h('p', { className: 'text-sm mb-5', style: { color: 'var(--text-dim)' } }, 'Current: ', h('span', { style: { color: 'var(--text)' } }, user.email || '(none)')),
        h('form', { onSubmit: handleEmailSave },
          field('New email address',
            h('input', { type: 'email', value: newEmail, onChange: e => setNewEmail(e.target.value), className: 'w-full', autoComplete: 'email' })),
          field('Confirm with your password',
            h('input', { type: 'password', value: emailPassword, onChange: e => setEmailPassword(e.target.value), className: 'w-full', autoComplete: 'current-password' })),
          h('button', { type: 'submit', disabled: emailSaving, className: 'gradient-btn px-6 py-2.5 text-white font-semibold rounded-xl text-sm disabled:opacity-50' },
            emailSaving ? 'Saving...' : 'Update email')
        )
      ),

      h('div', { className: 'card' },
        h('h3', { className: 'text-lg font-semibold text-white mb-5' }, 'Change password'),
        h('form', { onSubmit: handlePasswordSave },
          field('Current password',
            h('input', { type: 'password', value: curPassword, onChange: e => setCurPassword(e.target.value), className: 'w-full', autoComplete: 'current-password' })),
          field('New password (min. 6 characters)',
            h('input', { type: 'password', value: newPassword, onChange: e => setNewPassword(e.target.value), className: 'w-full', autoComplete: 'new-password' })),
          field('Confirm new password',
            h('input', { type: 'password', value: confirmPassword, onChange: e => setConfirmPassword(e.target.value), className: 'w-full', autoComplete: 'new-password' })),
          h('button', { type: 'submit', disabled: pwSaving, className: 'gradient-btn px-6 py-2.5 text-white font-semibold rounded-xl text-sm disabled:opacity-50' },
            pwSaving ? 'Saving...' : 'Change password')
        )
      )
    )
  );
}

// ====================== DASHBOARD ======================
function UserDashboard({ user, showToast, onLogout }) {
  const [projects, setProjects] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showCreate, setShowCreate] = useState(false);
  const [newName, setNewName] = useState('');
  const [creating, setCreating] = useState(false);
  const [renameId, setRenameId] = useState(null);
  const [renameValue, setRenameValue] = useState('');
  const [notesId, setNotesId] = useState(null);
  const [notesValue, setNotesValue] = useState('');

  const loadProjects = useCallback(async () => {
    try { setProjects(await api('get_projects')); } catch (e) { showToast(e.message, 'error'); }
    finally { setLoading(false); }
  }, [showToast]);

  useEffect(() => { loadProjects(); }, [loadProjects]);

  const handleCreate = async (e) => {
    e.preventDefault();
    if (!newName.trim()) return;
    setCreating(true);
    try {
      const p = await api('create_project', 'POST', { name: newName.trim() });
      showToast('Project created!');
      setNewName('');
      setShowCreate(false);
      window.location.hash = `#/project/${p.id}`;
    } catch (e2) { showToast(e2.message, 'error'); }
    finally { setCreating(false); }
  };

  const handleDelete = async (id) => {
    if (!confirm('Delete this project? This cannot be undone.')) return;
    try {
      await api('delete_project', 'POST', { id });
      showToast('Project deleted');
      loadProjects();
    } catch (e) { showToast(e.message, 'error'); }
  };

  const handleCopy = async (id) => {
    try {
      const p = await api('copy_project', 'POST', { project_id: id });
      showToast('Project copied: ' + p.name);
      loadProjects();
    } catch (e) { showToast(e.message, 'error'); }
  };

  const handleRename = async (id) => {
    if (!renameValue.trim()) { showToast('Name cannot be empty', 'error'); return; }
    try {
      await api('rename_project', 'POST', { id, name: renameValue.trim() });
      showToast('Project renamed');
      setRenameId(null);
      setRenameValue('');
      loadProjects();
    } catch (e) { showToast(e.message, 'error'); }
  };

  const handleNotesSave = async (id) => {
    try {
      await api('update_notes', 'POST', { project_id: id, notes: notesValue });
      showToast('Notes saved');
      setNotesId(null);
      loadProjects();
    } catch (e) { showToast(e.message, 'error'); }
  };

  const projList = Array.isArray(projects) ? projects : Object.values(projects);

  return h('div', { className: 'min-h-screen' },
    h(AppHeader, { user, onLogout, subtitle: 'Welcome, ' + user.username }),
    h('main', { className: 'max-w-6xl mx-auto p-6' },
      h('div', { className: 'flex items-center justify-between mb-6 mt-2' },
        h('div', null,
          h('h2', { className: 'text-2xl font-bold text-white' }, 'My Projects'),
          h('p', { className: 'text-sm mt-1', style: { color: 'var(--text-dim)' } }, projList.length + ' project' + (projList.length === 1 ? '' : 's'))
        ),
        h('button', { onClick: () => setShowCreate(true), className: 'gradient-btn px-5 py-2.5 text-white font-semibold rounded-xl text-sm' }, '+ New Project')
      ),
      showCreate && h('div', { className: 'card mb-6 animate-slide-up' },
        h('form', { onSubmit: handleCreate, className: 'flex gap-3 items-center' },
          h('input', { type: 'text', value: newName, onChange: e => setNewName(e.target.value), placeholder: 'my-awesome-site', className: 'flex-1', autoFocus: true }),
          h('button', { type: 'submit', disabled: creating, className: 'gradient-btn px-6 py-2.5 text-white font-semibold rounded-xl text-sm disabled:opacity-50' }, creating ? 'Creating...' : 'Create'),
          h('button', { type: 'button', onClick: () => setShowCreate(false), className: 'btn-ghost px-4 py-2.5 rounded-xl text-sm' }, 'Cancel')
        )
      ),
      loading ? h('div', { className: 'empty-state' },
        h('div', { className: 'spinner mx-auto mb-4' }),
        h('p', null, 'Loading projects...')
      ) : projList.length === 0 ? h('div', { className: 'card empty-state' },
        h('div', { className: 'icon' }, '\u{1F680}'),
        h('p', { className: 'text-lg mb-2', style: { color: 'var(--text)' } }, 'No projects yet'),
        h('p', { className: 'text-sm' }, 'Create your first project and let AI build your website.')
      ) : h('div', { className: 'grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-5' },
        projList.map(p => h('div', { key: p.id, className: 'card card-interactive', onClick: () => window.location.hash = `#/project/${p.id}` },
          h('div', { className: 'flex items-start justify-between mb-3' },
            h('div', { style: { width: 42, height: 42, borderRadius: 12, background: 'rgba(249,115,22,0.12)', border: '1px solid rgba(249,115,22,0.25)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 } }, '\u{1F310}'),
            h('div', { className: 'flex items-center gap-1 flex-wrap' },
              h('button', {
                className: 'btn-ghost rounded-lg text-xs px-2 py-1',
                title: 'Rename project',
                onClick: e => { e.stopPropagation(); setRenameId(p.id); setRenameValue(p.name); }
              }, 'Rename'),
              h('button', {
                className: 'btn-ghost rounded-lg text-xs px-2 py-1',
                title: 'Duplicate this project (website, photos and prompt history - backups are not copied)',
                onClick: e => { e.stopPropagation(); handleCopy(p.id); }
              }, 'Copy'),
              h('button', {
                className: 'btn-ghost rounded-lg text-xs px-2 py-1',
                title: 'Notes',
                onClick: e => { e.stopPropagation(); setNotesId(p.id); setNotesValue(p.notes || ''); }
              }, 'Notes'),
              h('button', {
                className: 'btn-ghost rounded-lg text-xs px-2 py-1',
                style: { color: 'var(--danger)' },
                onClick: e => { e.stopPropagation(); handleDelete(p.id); }
              }, 'Delete')
            )
          ),
          h('h3', { className: 'text-lg font-semibold text-white mb-1' }, p.name),
          h('p', { className: 'text-xs', style: { color: 'var(--text-mute)' } }, 'Updated ' + new Date(p.updated_at).toLocaleString()),
          p.notes && h('p', { className: 'text-xs mt-2', style: { color: 'var(--text-dim)', background: 'rgba(255,255,255,0.04)', padding: '6px 8px', borderRadius: 6, whiteSpace: 'pre-wrap', wordBreak: 'break-word' } }, p.notes.length > 120 ? p.notes.substring(0,120) + '...' : p.notes),
          h('div', { className: 'mt-4 flex items-center gap-2 text-sm', style: { color: 'var(--accent-2)' } },
            'Open workspace ', h('span', null, '\u2192')
          )
        ))
      ),
      renameId && h('div', { className: 'modal d-block', tabIndex: -1, role: 'dialog', onClick: () => setRenameId(null) },
        h('div', { className: 'modal-dialog', style: { maxWidth: 420, margin: '80px auto', padding: 12, display: 'flex', justifyContent: 'center' }, onClick: e => e.stopPropagation() },
          h('div', { className: 'modal-content', style: { width: '100%', padding: 20 } },
            h('h3', { className: 'text-sm font-bold text-white mb-3' }, 'Rename project'),
            h('input', { type: 'text', value: renameValue, onChange: e => setRenameValue(e.target.value), className: 'w-full mb-4', autoFocus: true }),
            h('div', { className: 'flex gap-2 justify-end' },
              h('button', { onClick: () => setRenameId(null), className: 'btn-ghost px-4 py-2 rounded-lg text-sm' }, 'Cancel'),
              h('button', { onClick: () => handleRename(renameId), className: 'gradient-btn px-5 py-2 text-white rounded-lg text-sm' }, 'Save')
            )
          )
        )
      ),
      notesId && h('div', { className: 'modal d-block', tabIndex: -1, role: 'dialog', onClick: () => setNotesId(null) },
        h('div', { className: 'modal-dialog', style: { maxWidth: 520, margin: '80px auto', padding: 12, display: 'flex', justifyContent: 'center' }, onClick: e => e.stopPropagation() },
          h('div', { className: 'modal-content', style: { width: '100%', padding: 20 } },
            h('h3', { className: 'text-sm font-bold text-white mb-3' }, 'Project notes'),
            h('textarea', { value: notesValue, onChange: e => setNotesValue(e.target.value), rows: 8, className: 'w-full mb-4 resize-y', placeholder: 'Add a long note for this project...', style: { minHeight: 120 } }),
            h('div', { className: 'flex gap-2 justify-end' },
              h('button', { onClick: () => setNotesId(null), className: 'btn-ghost px-4 py-2 rounded-lg text-sm' }, 'Cancel'),
              h('button', { onClick: () => handleNotesSave(notesId), className: 'gradient-btn px-5 py-2 text-white rounded-lg text-sm' }, 'Save notes')
            )
          )
        )
      )
    )
  );
}

// ====================== ADMIN ======================
function AdminPanel({ user, showToast, onLogout }) {
  const [tab, setTab] = useState('config');
  const [config, setConfig] = useState({});
  const [configLoading, setConfigLoading] = useState(true);
  const [configSaving, setConfigSaving] = useState(false);
  const [users, setUsers] = useState([]);
  const [usersLoading, setUsersLoading] = useState(false);
  const [models, setModels] = useState([]);
  const [modelsLoading, setModelsLoading] = useState(false);
  const [modelsError, setModelsError] = useState('');
  const [modelCustom, setModelCustom] = useState(false);

  const loadModels = useCallback(async (endpoint, apiKey) => {
    if (!endpoint) { setModels([]); setModelsError('Set the LLM endpoint URL first, then refresh.'); return; }
    setModelsLoading(true);
    setModelsError('');
    try {
      const data = await api('list_models', 'POST', { endpoint, api_key: apiKey || '' });
      const list = Array.isArray(data) ? data : Object.values(data || {});
      setModels(list);
      if (list.length === 0) setModelsError('The endpoint reported no models.');
    } catch (e) {
      setModels([]);
      setModelsError(e.message);
    } finally { setModelsLoading(false); }
  }, []);

  const loadConfig = useCallback(async () => {
    setConfigLoading(true);
    try {
      const cfg = await api('get_config');
      setConfig(cfg);
      if (cfg.llm_endpoint) loadModels(cfg.llm_endpoint, cfg.llm_api_key);
    } catch (e) { showToast(e.message, 'error'); }
    finally { setConfigLoading(false); }
  }, [showToast, loadModels]);

  const loadUsers = useCallback(async () => {
    setUsersLoading(true);
    try { setUsers(await api('admin_get_users')); } catch (e) { showToast(e.message, 'error'); }
    finally { setUsersLoading(false); }
  }, [showToast]);

  useEffect(() => { loadConfig(); }, [loadConfig]);
  useEffect(() => { if (tab === 'users') loadUsers(); }, [tab, loadUsers]);

  const handleSaveConfig = async (e) => {
    e.preventDefault();
    setConfigSaving(true);
    try {
      await api('save_config', 'POST', config);
      showToast('Configuration saved');
    } catch (err) { showToast(err.message, 'error'); }
    finally { setConfigSaving(false); }
  };

  const handleRoleChange = async (userId, newRole) => {
    try {
      await api('admin_update_user', 'POST', { user_id: userId, role: newRole });
      showToast('User role updated');
      loadUsers();
    } catch (e) { showToast(e.message, 'error'); }
  };

  const userList = Array.isArray(users) ? users : Object.values(users);

  const field = (label, hint, input) => h('div', { className: 'mb-5' },
    h('label', { className: 'section-label block mb-2' }, label),
    input,
    hint && h('p', { className: 'text-xs mt-1.5', style: { color: 'var(--text-mute)' } }, hint)
  );

  const txt = (key, placeholder, type) => h('input', {
    type: type || 'text',
    value: config[key] || '',
    onChange: e => setConfig({ ...config, [key]: e.target.value }),
    placeholder: placeholder || '',
    className: 'w-full'
  });

  const area = (key, rows, placeholder) => h('textarea', {
    rows: rows || 3,
    value: config[key] || '',
    onChange: e => setConfig({ ...config, [key]: e.target.value }),
    placeholder: placeholder || '',
    className: 'w-full resize-y'
  });

  const tabLabels = { config: '\u{1F9E0} LLM Configuration', email: '\u{1F4E7} Email settings', users: '\u{1F465} Users' };

  return h('div', { className: 'min-h-screen' },
    h(AppHeader, { user, onLogout, subtitle: 'Admin Panel' }),
    h('main', { className: 'max-w-4xl mx-auto p-6' },
      h('div', { className: 'flex gap-2 mb-6 mt-2' },
        ['config', 'email', 'users'].map(t => h('button', {
          key: t,
          onClick: () => setTab(t),
          className: `px-5 py-2.5 rounded-xl text-sm font-semibold transition-all ${tab === t ? 'gradient-btn text-white' : 'btn-ghost'}`
        }, tabLabels[t]))
      ),
      tab === 'config' && (configLoading ? h('div', { className: 'empty-state' }, h('div', { className: 'spinner mx-auto' }))
        : h('div', { className: 'card' },
          h('form', { onSubmit: handleSaveConfig },
            field('LLM Endpoint URL', 'Base URL of your LM Studio / Ollama server',
              txt('llm_endpoint', 'http://localhost:1234')),
            field('API Key', 'Required for LM Studio remote access',
              txt('llm_api_key', 'sk-lm-...', 'password')),
            field('Model Name', modelsError || (models.length ? models.length + ' model(s) reported by the endpoint' : null),
              h('div', { className: 'flex gap-2' },
                (models.length > 0 && !modelCustom)
                  ? h('select', {
                      value: config.llm_model || '',
                      onChange: e => setConfig({ ...config, llm_model: e.target.value }),
                      className: 'flex-1'
                    },
                      h('option', { value: '' }, '\u2014 select a model \u2014'),
                      // keep the saved model visible even if the endpoint no longer lists it
                      (config.llm_model && !models.includes(config.llm_model))
                        ? h('option', { value: config.llm_model }, config.llm_model + ' (saved, not reported)')
                        : null,
                      models.map(m => h('option', { key: m, value: m }, m))
                    )
                  : h('input', {
                      type: 'text',
                      value: config.llm_model || '',
                      onChange: e => setConfig({ ...config, llm_model: e.target.value }),
                      placeholder: 'e.g. claude-sonnet-4-6 or google/gemma-4-e4b',
                      className: 'flex-1'
                    }),
                models.length > 0 && h('button', {
                  type: 'button',
                  onClick: () => setModelCustom(c => !c),
                  title: modelCustom ? 'Choose from the model list' : 'Type a model name manually',
                  className: 'btn-ghost px-4 rounded-lg text-sm',
                  style: { flexShrink: 0 }
                }, modelCustom ? '\u2630 List' : '\u270E Edit'),
                h('button', {
                  type: 'button',
                  onClick: () => loadModels(config.llm_endpoint, config.llm_api_key),
                  disabled: modelsLoading,
                  title: 'Fetch the available models from the endpoint',
                  className: 'btn-ghost px-4 rounded-lg text-sm disabled:opacity-50',
                  style: { flexShrink: 0 }
                }, modelsLoading ? h('div', { className: 'spinner spinner-sm' }) : '\u21BB Models')
              )),
            h('div', { className: 'mb-5 p-4 rounded-xl', style: { background: 'rgba(59,130,246,0.06)', border: '1px solid rgba(59,130,246,0.2)' } },
              h('label', { className: 'flex items-center gap-3 cursor-pointer' },
                h('input', { type: 'checkbox', checked: config.enable_wikipedia === '1', onChange: e => setConfig({ ...config, enable_wikipedia: e.target.checked ? '1' : '0' }), style: { width: 17, height: 17, accentColor: 'var(--accent)' } }),
                h('div', null,
                  h('span', { className: 'text-sm font-semibold text-white block' }, 'Wikipedia tools'),
                  h('span', { className: 'text-xs', style: { color: 'var(--text-dim)' } }, 'AI can research topics via search_wikipedia & get_wikipedia_page')
                )
              )
            ),
            field('System Prompt (optional)', null,
              area('system_prompt', 3, 'Extra instructions for the AI...')),
            field('Extra LLM Parameters (JSON, optional)', 'e.g. "temperature": 0.7, "max_tokens": 8192',
              h('textarea', { rows: 2, value: config.llm_parameters || '', onChange: e => setConfig({ ...config, llm_parameters: e.target.value }), placeholder: '"temperature": 0.7', className: 'w-full resize-y', style: { fontFamily: 'monospace', fontSize: 13 } })),
            field('Internal URL (async worker)', 'The address where this app reaches ITSELF, bypassing any reverse proxy. Must match the port ASPPY listens on (e.g. http://127.0.0.1:8058 behind IIS ARR).',
              txt('internal_url', 'http://127.0.0.1:8080')),
            h('button', { type: 'submit', disabled: configSaving, className: 'gradient-btn px-7 py-3 text-white font-semibold rounded-xl disabled:opacity-50' },
              configSaving ? 'Saving...' : 'Save Configuration')
          )
        )),
      tab === 'email' && (configLoading ? h('div', { className: 'empty-state' }, h('div', { className: 'spinner mx-auto' }))
        : h('div', { className: 'card' },
          h('form', { onSubmit: handleSaveConfig },
            h('h3', { className: 'text-lg font-semibold text-white mb-4' }, 'SMTP server'),
            h('div', { className: 'grid grid-cols-1 md:grid-cols-2 gap-x-5' },
              field('SMTP host', null, txt('smtp_host', 'smtp.example.com')),
              field('SMTP port', '25, 465 (SSL) or 587 (STARTTLS)', txt('smtp_port', '587')),
              field('SMTP login', null, txt('smtp_user', 'user@example.com')),
              field('SMTP password', null, txt('smtp_password', '', 'password'))
            ),
            field('From address', 'Sender of all outgoing emails (defaults to the SMTP login)', txt('smtp_from', 'noreply@example.com')),
            h('div', { className: 'mb-5 p-4 rounded-xl', style: { background: 'rgba(59,130,246,0.06)', border: '1px solid rgba(59,130,246,0.2)' } },
              h('label', { className: 'flex items-center gap-3 cursor-pointer' },
                h('input', { type: 'checkbox', checked: config.smtp_ssl === '1', onChange: e => setConfig({ ...config, smtp_ssl: e.target.checked ? '1' : '0' }), style: { width: 17, height: 17, accentColor: 'var(--accent)' } }),
                h('div', null,
                  h('span', { className: 'text-sm font-semibold text-white block' }, 'Use SSL/TLS'),
                  h('span', { className: 'text-xs', style: { color: 'var(--text-dim)' } }, 'SMTPS on port 465, STARTTLS on other ports')
                )
              )
            ),
            field('Public URL of this app', 'Used to build links in emails, e.g. the password reset link',
              txt('public_url', 'http://localhost:8080')),

            h('h3', { className: 'text-lg font-semibold text-white mb-4 mt-8' }, 'Password reset email'),
            field('Subject', null, txt('reset_email_subject', 'Reset your Autonome password')),
            field('Body', 'Placeholders: {username} and {link} (the reset link)', area('reset_email_body', 7)),

            h('h3', { className: 'text-lg font-semibold text-white mb-4 mt-8' }, 'Generation finished email'),
            field('Subject', null, txt('notify_email_subject', 'Your website generation has finished')),
            field('Body', 'Placeholders: {username}, {project}, {status} and {link} (the app URL)', area('notify_email_body', 6)),

            h('button', { type: 'submit', disabled: configSaving, className: 'gradient-btn px-7 py-3 text-white font-semibold rounded-xl disabled:opacity-50' },
              configSaving ? 'Saving...' : 'Save Email Settings')
          )
        )),
      tab === 'users' && (usersLoading ? h('div', { className: 'empty-state' }, h('div', { className: 'spinner mx-auto' }))
        : h('div', { className: 'card', style: { padding: 0, overflow: 'hidden' } },
          h('table', { className: 'w-full text-sm' },
            h('thead', null,
              h('tr', { style: { borderBottom: '1px solid var(--border)', background: 'rgba(16,24,39,0.5)' } },
                ['User', 'Email', 'Role', 'Created', 'Change role'].map(hd => h('th', { key: hd, className: 'text-left py-3.5 px-5', style: { color: 'var(--text-dim)' } }, hd))
              )
            ),
            h('tbody', null,
              userList.map(u => h('tr', { key: u.id, style: { borderBottom: '1px solid var(--border)' } },
                h('td', { className: 'py-3.5 px-5 text-white font-medium' }, u.username),
                h('td', { className: 'py-3.5 px-5 text-xs', style: { color: 'var(--text-dim)' } }, u.email || '\u2014'),
                h('td', { className: 'py-3.5 px-5' }, h('span', { className: `badge ${u.role === 'admin' ? 'badge-admin' : 'badge-user'}` }, u.role)),
                h('td', { className: 'py-3.5 px-5 text-xs', style: { color: 'var(--text-dim)' } }, new Date(u.created_at).toLocaleDateString()),
                h('td', { className: 'py-3.5 px-5' },
                  h('select', { value: u.role, onChange: e => handleRoleChange(u.id, e.target.value), className: 'text-xs py-1.5 px-3' },
                    h('option', { value: 'user' }, 'User'),
                    h('option', { value: 'admin' }, 'Admin')
                  )
                )
              ))
            )
          )
        ))
    )
  );
}

// ====================== HELPERS: copy, duration ======================
function copyToClipboard(text, showToast) {
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).then(() => showToast && showToast('Copied to clipboard')).catch(() => fallbackCopy(text, showToast));
  } else {
    fallbackCopy(text, showToast);
  }
}
function fallbackCopy(text, showToast) {
  const ta = document.createElement('textarea');
  ta.value = text;
  ta.style.position = 'fixed';
  ta.style.opacity = '0';
  document.body.appendChild(ta);
  ta.select();
  try { document.execCommand('copy'); showToast && showToast('Copied to clipboard'); } catch(e) { showToast && showToast('Copy failed', 'error'); }
  document.body.removeChild(ta);
}
function formatDuration(secs) {
  if (!secs || secs <= 0) return '';
  if (secs >= 60) return Math.floor(secs/60) + 'm ' + (secs%60) + 's';
  return secs + 's';
}
function formatChatTime(iso) {
  if (!iso) return '';
  try {
    const d = new Date(iso);
    if (isNaN(d.getTime())) return iso;
    return d.toLocaleString();
  } catch(e) { return iso; }
}

// ====================== USER MESSAGE (with copy) ======================
function UserMessage({ msg, showToast }) {
  const content = msg.content || '';
  const created = msg.created_at || '';
  const handleCopy = () => copyToClipboard(content, showToast);
  return h('div', { className: 'chat-bubble chat-bubble-user', style: { position: 'relative' } },
    h('div', null, content),
    h('div', { className: 'chat-meta' },
      created && h('span', { style: { opacity: 0.85 } }, formatChatTime(created)),
      h('button', { onClick: handleCopy, className: 'copy-btn', title: 'Copy prompt' }, '\u2398 Copy')
    )
  );
}

// ====================== AI MESSAGE (with reasoning) ======================
function AiMessage({ msg, showToast }) {
  const content = msg.content || '';
  const reasoning = msg.reasoning || '';
  const dur = msg.duration_seconds || 0;
  const tokIn = msg.tokens_in || 0;
  const tokOut = msg.tokens_out || 0;
  const created = msg.created_at || '';
  let tps = '';
  if (dur > 0 && tokOut > 0) {
    const v = Math.round((tokOut / dur) * 10) / 10;
    tps = v.toFixed(1).replace('.', ',') + ' t/s';
  }
  return h('div', { className: 'chat-bubble chat-bubble-ai', style: { position: 'relative' } },
    reasoning && h('div', null,
      h('span', { className: 'reasoning-label' }, '\u{1F9E0} thinking'),
      h('div', { className: 'reasoning-block' }, reasoning)
    ),
    content && h('div', null, content.length > 1200 ? content.substring(0, 1200) + '\u2026' : content),
    (dur > 0 || tokOut > 0 || created) && h('div', { className: 'chat-meta chat-meta-ai' },
      dur > 0 && h('span', null, '\u23F1 ' + formatDuration(dur)),
      tps && h('span', null, tps),
      tokOut > 0 && h('span', null, tokOut + ' out' + (tokIn ? ' / ' + tokIn + ' in' : '')),
      created && h('span', null, formatChatTime(created))
    )
  );
}

// ====================== SUGGESTIONS ======================
const FIRST_PROMPT_SUGGESTIONS = [
  "Surprise me with a mindblowing website about a topic of your choice \u2014 make it visually bold and unexpected.",
  "Create a dark-themed, animated website about traveling in Australia, with smooth scroll animations and a hero section that feels alive.",
  "Design a retro-futuristic 80s synthwave landing page for a fictional space travel agency, with neon gradients and glowing text effects.",
  "Build a minimalist one-page portfolio for a fictional architect, with elegant typography, generous whitespace, and subtle hover animations.",
  "Make an interactive, playful website for a fictional coffee brand, with a parallax scroll effect and fun micro-interactions on buttons and cards."
];

const IMPROVEMENT_GROUPS = [
  { label: "Tone Of Voice", items: [
    "Make all texts more friendly and conversational.",
    "Make all texts warmer and more inviting.",
    "Make all texts more formal and authoritative.",
    "Make all texts punchier and more energetic, with shorter sentences."
  ]},
  { label: "Theme & Atmosphere", items: [
    "Switch to a dark theme with high contrast.",
    "Make it a light, airy theme with soft pastel colors.",
    "Give it a more corporate/professional look.",
    "Make the overall vibe more playful and colorful."
  ]},
  { label: "Layout & Structure", items: [
    "Make the hero section full-screen.",
    "Add more whitespace, it feels too cramped."
  ]},
  { label: "Color & Branding", items: [
    "Change the color palette to shades of blue and teal.",
    "Make the accent color more vibrant / eye-catching.",
    "Tone down the colors, make it more minimalist."
  ]},
  { label: "Content Structure", items: [
    "Add a testimonials section.",
    "Add a call-to-action button at the end of every section.",
    "Reorganize the content so the most important info comes first."
  ]}
];


function SuggestionsModal({ isFirst, onSelect, onClose }) {
  const handleSelect = (text) => {
    onSelect(text);
  };
  return h('div', { className: 'modal d-block', tabIndex: -1, role: 'dialog', 'aria-modal': true, onClick: onClose },
    h('div', { className: 'modal-dialog', style: { maxWidth: '560px', margin: '60px auto', padding: '12px', height: 'auto', display: 'flex', justifyContent: 'center' }, onClick: e => e.stopPropagation() },
      h('div', { className: 'modal-content', style: { width: '100%' } },
        h('div', { className: 'modal-header' },
          h('h5', { className: 'modal-title' }, isFirst ? '\u2728 Starter prompts' : '\uD83D\uDCA1 Improve your website'),
          h('button', { type: 'button', className: 'btn-close', 'aria-label': 'Close', onClick: onClose }, '\u00D7')
        ),
        h('div', { className: 'modal-body', style: { padding: '16px', overflowY: 'auto', maxHeight: '65vh', flex: 'none' } },
          isFirst
            ? h('div', { style: { display: 'flex', flexDirection: 'column', gap: '8px' } },
                h('p', { className: 'text-xs mb-2', style: { color: 'var(--text-dim)' } }, 'Click a suggestion to add it to your prompt.'),
                FIRST_PROMPT_SUGGESTIONS.map((s, i) => h('button', {
                  key: i,
                  type: 'button',
                  onClick: () => handleSelect(s),
                  className: 'suggestion-item',
                  style: { textAlign: 'left' }
                }, s))
              )
            : h('div', { style: { display: 'flex', flexDirection: 'column', gap: '16px' } },
                h('p', { className: 'text-xs', style: { color: 'var(--text-dim)' } }, 'Click a suggestion to append it to your prompt.'),
                IMPROVEMENT_GROUPS.map(g => h('div', { key: g.label },
                  h('div', { className: 'section-label', style: { marginBottom: '6px', fontSize: '11px' } }, g.label),
                  h('div', { style: { display: 'flex', flexDirection: 'column', gap: '6px' } },
                    g.items.map((s, idx) => h('button', {
                      key: idx,
                      type: 'button',
                      onClick: () => handleSelect(s),
                      className: 'suggestion-item suggestion-item-sm'
                    }, s))
                  )
                ))
              )
        ),
        h('div', { className: 'modal-footer' },
          h('button', { type: 'button', onClick: onClose, className: 'btn-ghost px-5 py-2 rounded-xl text-sm' }, 'Close')
        )
      )
    )
  );
}

// ====================== BACKUP PREVIEW MODAL (full-width, Bootstrap style) ======================
function BackupPreviewModal({ projectId, backupName, onClose, onRestore, restoring }) {
  const src = `api.asp?action=preview_backup&project_id=${projectId}&backup=${encodeURIComponent(backupName)}`;
  const label = backupName.replace('T', ' \u00B7 ').replace(/-(\d\d)-(\d\d)$/, ':$1:$2');
  return h('div', { className: 'modal d-block', tabIndex: -1, role: 'dialog', 'aria-modal': true },
    h('div', { className: 'modal-dialog modal-fullscreen' },
      h('div', { className: 'modal-content' },
        h('div', { className: 'modal-header' },
          h('h5', { className: 'modal-title' }, '\u23F1\uFE0F Backup preview \u2014 ', h('span', { style: { fontFamily: 'Consolas, monospace', color: 'var(--accent-2)' } }, label)),
          h('button', { type: 'button', className: 'btn-close', 'aria-label': 'Close', onClick: onClose }, '\u00D7')
        ),
        h('div', { className: 'modal-body' },
          h('iframe', { src, className: 'backup-preview-frame', title: 'Backup preview' })
        ),
        h('div', { className: 'modal-footer' },
          h('span', { className: 'text-xs mr-auto', style: { color: 'var(--text-mute)' } }, 'Restoring backs up the current version first, so nothing is lost.'),
          h('button', { type: 'button', onClick: onClose, className: 'btn-ghost px-5 py-2.5 rounded-xl text-sm' }, 'Close'),
          h('button', {
            type: 'button',
            onClick: onRestore,
            disabled: restoring,
            className: 'gradient-btn px-6 py-2.5 text-white text-sm font-semibold rounded-xl disabled:opacity-50'
          }, restoring ? 'Restoring...' : '\u21A9 Restore this version')
        )
      )
    )
  );
}

// ====================== PROJECT WORKSPACE ======================
function ProjectWorkspace({ user, projectId, showToast, uploadImageRaw }) {
  const [project, setProject] = useState(null);
  const [images, setImages] = useState([]);
  const [backups, setBackups] = useState([]);
  const [showHistory, setShowHistory] = useState(false);
  const [restoring, setRestoring] = useState(false);
  const [previewBackup, setPreviewBackup] = useState(null);
  const [messages, setMessages] = useState([]);
  const [prompt, setPrompt] = useState('');
  const [generating, setGenerating] = useState(false);
  const [jobId, setJobId] = useState(null);
  const [notifyEmail, setNotifyEmail] = useState(false);
  const [loading, setLoading] = useState(true);
  const [uploading, setUploading] = useState(false);
  const [previewKey, setPreviewKey] = useState(0);
  const [showSuggestions, setShowSuggestions] = useState(false);
  const [showMobileMenu, setShowMobileMenu] = useState(false);
  const chatEndRef = useRef(null);
  const cancelledRef = useRef(false);
  const promptRef = useRef(null);

  const siteUrl = project && project.folder ? `/sites/${project.folder}/index.html` : '';

  const loadProject = useCallback(async () => {
    setLoading(true);
    try {
      const projects = await api('get_projects');
      const projList = Array.isArray(projects) ? projects : Object.values(projects);
      setProject(projList.find(p => String(p.id) === String(projectId)));
    } catch (e) { showToast(e.message, 'error'); }
    finally { setLoading(false); }
  }, [projectId, showToast]);

  const loadImages = useCallback(async () => {
    try {
      const imgs = await api(`list_images&project_id=${projectId}`);
      setImages(Array.isArray(imgs) ? imgs : Object.values(imgs));
    } catch (e) { /* ignore */ }
  }, [projectId]);

  const loadBackups = useCallback(async () => {
    try {
      const b = await api(`list_backups&project_id=${projectId}`);
      const list = Array.isArray(b) ? b : Object.values(b);
      list.sort((a, x) => x.name.localeCompare(a.name)); // newest first
      setBackups(list);
    } catch (e) { /* ignore */ }
  }, [projectId]);

  const loadChatHistory = useCallback(async () => {
    try {
      const hist = await api(`get_chat_history&project_id=${projectId}`);
      const list = Array.isArray(hist) ? hist : Object.values(hist || {});
      // Ensure messages have role/content structure expected by UI
      setMessages(list);
    } catch (e) { /* ignore - will show empty */ }
  }, [projectId]);

  useEffect(() => { loadProject(); loadImages(); loadBackups(); loadChatHistory(); }, [loadProject, loadImages, loadBackups, loadChatHistory]);
  useEffect(() => { chatEndRef.current?.scrollIntoView({ behavior: 'smooth' }); }, [messages, generating]);

  // Keep session alive while in workspace
  useEffect(() => {
    const iv = setInterval(() => api('keepalive').catch(()=>{}), 5*60*1000);
    return () => clearInterval(iv);
  }, []);

  const refreshAll = () => {
    loadImages();
    loadBackups();
    setPreviewKey(k => k + 1);
  };

  // Poll a background job until it finishes (or is cancelled)
  const pollJob = async (jid) => {
    const maxPolls = 1800; // 4s x 1800 = 2 hours
    for (let p = 0; p < maxPolls; p++) {
      await sleep(4000);
      if (cancelledRef.current) return { status: 'cancelled' };
      let st;
      try {
        st = await api(`generate_status&job_id=${jid}`);
      } catch (pollErr) {
        continue; // transient network error: keep polling
      }
      if (st.status === 'done') return { status: 'done', result: st.result || {} };
      if (st.status === 'cancelled') return { status: 'cancelled' };
      if (st.status === 'error') {
        const msg = (st.result && st.result.error) ? st.result.error : 'Generation failed';
        return { status: 'error', error: msg };
      }
    }
    return { status: 'timeout' };
  };

  const finishJob = (outcome) => {
    setJobId(null);
    setGenerating(false);
    if (outcome.status === 'done') {
      const data = outcome.result || {};
      if (data.message) {
        const assistantMsg = { role: 'assistant', reasoning: data.reasoning || '', content: '\u2705 ' + data.message, tokens_in: data.tokens_in || 0, tokens_out: data.tokens_out || 0, duration_seconds: data.duration_seconds || 0, created_at: new Date().toISOString() };
        setMessages(prev => [...prev, assistantMsg]);
        showToast('Website generated!');
        refreshAll();
        loadChatHistory();
      } else if (data.parsed === false && data.raw) {
        setMessages(prev => [...prev, { role: 'assistant', reasoning: data.reasoning || '', content: data.raw, created_at: new Date().toISOString() }]);
        showToast('AI responded, but not in file format. Try again.', 'error');
        loadChatHistory();
      } else {
        setMessages(prev => [...prev, { role: 'assistant', content: 'Empty response. Try again.', created_at: new Date().toISOString() }]);
        showToast('Empty response from AI', 'error');
      }
    } else if (outcome.status === 'cancelled') {
      setMessages(prev => [...prev, { role: 'assistant', content: '\u{1F6AB} Generation cancelled. Nothing was changed.', created_at: new Date().toISOString() }]);
      showToast('Generation cancelled');
      loadChatHistory();
    } else if (outcome.status === 'error') {
      setMessages(prev => [...prev, { role: 'assistant', content: '\u274C ' + outcome.error, created_at: new Date().toISOString() }]);
      showToast(outcome.error, 'error');
      loadChatHistory();
    } else {
      setMessages(prev => [...prev, { role: 'assistant', content: '\u274C Generation timed out after 2 hours.', created_at: new Date().toISOString() }]);
      showToast('Generation timed out', 'error');
    }
  };

  // On load: pick up a generation that is still running in the background
  useEffect(() => {
    let alive = true;
    (async () => {
      try {
        const job = await api(`active_job&project_id=${projectId}`);
        if (!alive || !job || !job.job_id) return;
        cancelledRef.current = false;
        setJobId(job.job_id);
        setNotifyEmail(job.notify === 1);
        setGenerating(true);
        setMessages(prev => [...prev,
          { role: 'user', content: job.prompt },
          { role: 'assistant', content: '\u23F3 A generation for this project is still running in the background. You cannot submit a new prompt until it finishes (or you cancel it).' }
        ]);
        const outcome = await pollJob(job.job_id);
        if (alive) finishJob(outcome);
      } catch (e) { /* no active job */ }
    })();
    return () => { alive = false; };
  }, [projectId]); // eslint-disable-line

  const handleReset = async () => {
    if (generating) { showToast('Wait for the running generation to finish (or cancel it) first.', 'error'); return; }
    if (!confirm('Start all over again? The website, prompt history and backups will be removed.\n\nYour uploaded photos are kept.')) return;
    try {
      const msg = await api('reset_project', 'POST', { project_id: parseInt(projectId) });
      setMessages([]);
      setShowHistory(false);
      setPreviewBackup(null);
      showToast(typeof msg === 'string' ? msg : 'Project reset');
      refreshAll();
    } catch (e) { showToast(e.message, 'error'); }
  };

  const handleRestore = async (backupName) => {
    if (!confirm('Restore version ' + backupName + '? The current version is backed up first.')) return;
    setRestoring(true);
    try {
      const data = await api('restore_backup', 'POST', { project_id: parseInt(projectId), backup: backupName });
      showToast(data.message);
      setMessages(prev => [...prev, { role: 'assistant', content: '\u23F1\uFE0F ' + data.message }]);
      setShowHistory(false);
      setPreviewBackup(null);
      refreshAll();
    } catch (e) { showToast(e.message, 'error'); }
    finally { setRestoring(false); }
  };

  const handleDeleteImage = async (filename) => {
    if (!confirm('Delete ' + filename + '?')) return;
    try {
      await api('delete_image', 'POST', { project_id: parseInt(projectId), filename });
      showToast('Image deleted');
      loadImages();
    } catch (e) { showToast(e.message, 'error'); }
  };

  const handleGenerate = async (e) => {
    e?.preventDefault();
    if (!prompt.trim() || generating) return;
    const userMsg = prompt.trim();
    setPrompt('');
    cancelledRef.current = false;
    setMessages(prev => [...prev, { role: 'user', content: userMsg }]);
    setGenerating(true);
    try {
      const data = await api('generate_website', 'POST', { project_id: parseInt(projectId), prompt: userMsg, notify: notifyEmail });

      if (data.job_id) {
        setJobId(data.job_id);
        const outcome = await pollJob(data.job_id);
        finishJob(outcome);
      } else {
        // Synchronous fallback: the response already contains the result
        finishJob({ status: 'done', result: data });
      }
    } catch (err) {
      setJobId(null);
      setGenerating(false);
      setMessages(prev => [...prev, { role: 'assistant', content: '\u274C ' + err.message }]);
      showToast(err.message, 'error');
    }
  };

  const handleCancel = async () => {
    if (!jobId) return;
    if (!confirm('Cancel this generation? The AI result will be discarded.')) return;
    try {
      await api('cancel_generation', 'POST', { job_id: jobId });
      cancelledRef.current = true; // the poll loop picks this up and finishes as cancelled
    } catch (e) { showToast(e.message, 'error'); }
  };

  const handlePromptKeyDown = (e) => {
    // Enter submits; Shift+Enter inserts a newline
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleGenerate();
    }
  };

  const handleUpload = async (e) => {
    const fl = Array.from(e.target.files || []);
    if (fl.length === 0) return;
    const allowed = ['image/jpeg', 'image/png', 'image/gif', 'image/webp', 'image/bmp', 'image/avif'];
    const invalid = fl.filter(f => !allowed.includes(f.type));
    if (invalid.length > 0) { showToast('Invalid format: ' + invalid.map(f => f.name).join(', '), 'error'); return; }
    setUploading(true);
    try {
      const data = await uploadImageRaw(fl, parseInt(projectId));
      const urls = data.urls ? Object.values(data.urls) : [data.url];
      setMessages(prev => [...prev, { role: 'assistant', content: '\u{1F4F7} ' + urls.length + ' image(s) uploaded:\n' + urls.join('\n') }]);
      showToast(urls.length + ' image(s) uploaded!');
      loadImages();
    } catch (err) { showToast(err.message, 'error'); }
    finally { setUploading(false); }
    e.target.value = '';
  };

  if (loading) {
    return h('div', { className: 'min-h-screen flex items-center justify-center' },
      h('div', { className: 'text-center' }, h('div', { className: 'spinner mx-auto mb-4' }), h('p', { style: { color: 'var(--text-dim)' } }, 'Loading project...'))
    );
  }

  if (!project) {
    return h('div', { className: 'min-h-screen flex items-center justify-center' },
      h('div', { className: 'card text-center' },
        h('p', { className: 'mb-4', style: { color: 'var(--text-dim)' } }, 'Project not found'),
        h('button', { onClick: () => window.location.hash = '#/dashboard', className: 'gradient-btn px-5 py-2.5 text-white rounded-xl' }, 'Back to Dashboard')
      )
    );
  }

  return h('div', { className: 'flex flex-col', style: { height: '100vh' } },
    // Header — desktop shows all actions inline, mobile collapses into hamburger
    h('header', { className: 'gradient-header project-header px-5 py-2.5', style: { position: 'relative' } },
      h('div', { className: 'flex items-center justify-between gap-2' },
        h('div', { className: 'flex items-center gap-3 min-w-0' },
          h('button', { onClick: () => window.location.hash = '#/dashboard', className: 'btn-ghost px-3 py-2 rounded-lg text-sm flex-shrink-0' }, '\u2190'),
          h('div', { className: 'app-logo', style: { width: 32, height: 32, fontSize: 16, borderRadius: 9, flexShrink: 0 } }, '\u{1F310}'),
          h('div', { className: 'min-w-0' },
            h('h1', { className: 'font-bold text-white project-title truncate', style: { fontSize: 15, lineHeight: 1.2 } }, project.name),
            h('span', { className: 'text-xs project-folder', style: { color: 'var(--text-mute)' } }, 'sites/' + (project.folder || ''))
          )
        ),
        h('div', { className: 'flex items-center gap-2 flex-shrink-0', style: { position: 'relative' } },
          // Desktop actions — hidden on mobile via CSS
          h('div', { className: 'desktop-actions flex items-center gap-2' },
            h('button', {
              onClick: handleReset,
              disabled: generating,
              title: 'Start all over again - removes the website, history and backups. Photos are kept.',
              className: 'btn-ghost px-3 py-2 rounded-lg text-sm disabled:opacity-50',
              style: { color: 'var(--danger)' }
            }, '\u21BA Reset'),
            h('button', { onClick: () => setShowHistory(!showHistory), className: 'btn-ghost px-3 py-2 rounded-lg text-sm' }, '\u23F1\uFE0F History' + (backups.length ? ' (' + backups.length + ')' : '')),
            h('a', { href: siteUrl, target: '_blank', className: 'btn-ghost px-3 py-2 rounded-lg text-sm no-underline', style: { textDecoration: 'none' } }, '\u{1F517} Open site'),
            h('a', { href: `api.asp?action=download_project&project_id=${projectId}`, className: 'btn-ghost px-3 py-2 rounded-lg text-sm no-underline', style: { textDecoration: 'none' }, title: 'Download the site as autonome.zip' }, '\u2B07 Download'),
            h('label', { className: 'gradient-btn-secondary px-4 py-2 text-white text-sm font-semibold rounded-lg cursor-pointer flex items-center gap-2' },
              uploading ? h('div', { className: 'spinner spinner-sm' }) : h('span', null, '\u{1F4F7}'),
              'Upload images',
              h('input', { type: 'file', accept: 'image/*', onChange: handleUpload, className: 'hidden', multiple: true })
            )
          ),
          // Mobile: Upload icon + hamburger
          h('div', { className: 'mobile-actions flex items-center gap-2' },
            h('label', { className: 'mobile-upload-btn gradient-btn-secondary text-white rounded-lg cursor-pointer flex items-center justify-center', style: { width: 38, height: 38 } },
              uploading ? h('div', { className: 'spinner spinner-sm' }) : h('span', { style: { fontSize: 18 } }, '\u{1F4F7}'),
              h('input', { type: 'file', accept: 'image/*', onChange: handleUpload, className: 'hidden', multiple: true })
            ),
            h('button', {
              onClick: () => setShowMobileMenu(!showMobileMenu),
              className: 'mobile-menu-btn btn-ghost rounded-lg flex items-center justify-center',
              style: { width: 38, height: 38, fontSize: 20, display: 'none' },
              'aria-label': 'Menu'
            }, showMobileMenu ? '\u2715' : '\u2630')
          ),
          // Mobile dropdown — same actions as desktop but stacked
          showMobileMenu && h('div', { className: 'mobile-dropdown' },
            h('button', { className: 'mobile-dropdown-item mobile-dropdown-item--danger', onClick: () => { setShowMobileMenu(false); handleReset(); } }, h('span', null, '\u21BA'), ' Reset project'),
            h('button', { className: 'mobile-dropdown-item', onClick: () => { setShowMobileMenu(false); setShowHistory(!showHistory); } }, h('span', null, '\u23F1\uFE0F'), ' History' + (backups.length ? ' (' + backups.length + ')' : '')),
            h('a', { href: siteUrl, target: '_blank', className: 'mobile-dropdown-item', style: { textDecoration: 'none' }, onClick: () => setShowMobileMenu(false) }, h('span', null, '\u{1F517}'), ' Open site'),
            h('a', { href: `api.asp?action=download_project&project_id=${projectId}`, className: 'mobile-dropdown-item', style: { textDecoration: 'none' }, onClick: () => setShowMobileMenu(false) }, h('span', null, '\u2B07'), ' Download zip'),
            h('label', { className: 'mobile-dropdown-item mobile-dropdown-item--primary', style: { cursor: 'pointer' } },
              h('span', null, '\u{1F4F7}'), ' Upload images',
              h('input', { type: 'file', accept: 'image/*', onChange: (e) => { setShowMobileMenu(false); handleUpload(e); }, className: 'hidden', multiple: true })
            )
          ),
          // History panel (desktop + mobile share same panel, positioned under header)
          showHistory && h('div', { className: 'history-panel' },
            h('div', { className: 'px-4 py-3', style: { borderBottom: '1px solid var(--border)' } },
              h('span', { className: 'text-sm font-bold text-white' }, 'Version history'),
              h('p', { className: 'text-xs', style: { color: 'var(--text-mute)' } }, 'Preview a version full-screen, or restore it directly')
            ),
            backups.length === 0
              ? h('div', { className: 'px-4 py-6 text-center text-sm', style: { color: 'var(--text-mute)' } }, 'No backups yet. A backup is created before every generation.')
              : h('div', { style: { maxHeight: 320, overflowY: 'auto' } },
                  backups.map(b => h('div', { key: b.name, className: 'history-row' },
                    h('div', null,
                      h('div', { className: 'text-sm text-white', style: { fontFamily: 'Consolas, monospace' } }, b.name.replace('T', ' \u00B7 ').replace(/-(\d\d)-(\d\d)$/, ':$1:$2')),
                      h('div', { className: 'text-xs', style: { color: 'var(--text-mute)' } }, b.files + ' file(s)')
                    ),
                    h('div', { className: 'flex items-center gap-2' },
                      h('button', {
                        onClick: () => { setPreviewBackup(b.name); setShowHistory(false); },
                        className: 'btn-ghost px-3 py-1.5 text-xs rounded-lg'
                      }, '\u{1F441}\uFE0F Preview'),
                      h('button', {
                        onClick: () => handleRestore(b.name),
                        disabled: restoring,
                        className: 'gradient-btn px-3 py-1.5 text-white text-xs font-semibold rounded-lg disabled:opacity-50'
                      }, restoring ? '...' : 'Restore')
                    )
                  ))
                )
          )
        )
      ),
      // Mobile backdrop to close menu when tapping outside
      showMobileMenu && h('div', { onClick: () => setShowMobileMenu(false), style: { position: 'fixed', inset: 0, zIndex: 45, background: 'transparent' } })
    ),

    // Body
    h('div', { className: 'flex-1 flex overflow-hidden workspace-body' },
      // Left: preview / code
      h('div', { className: 'flex-1 flex flex-col overflow-hidden', style: { background: 'var(--bg-soft)' } },
        // Images strip
        images.length > 0 && h('div', { className: 'px-5 pt-3 pb-1' },
          h('div', { className: 'flex items-center gap-2 mb-2' },
            h('span', { className: 'section-label' }, 'Images'),
            h('span', { className: 'text-xs', style: { color: 'var(--text-mute)' } }, images.length)
          ),
          h('div', { className: 'flex gap-3 overflow-x-auto pb-2' },
            images.map(img => h('div', { key: img.filename, className: 'img-thumb' },
              h('img', { src: img.url, alt: img.filename, title: img.filename }),
              h('button', { className: 'img-del', onClick: () => handleDeleteImage(img.filename), title: 'Delete' }, '\u00D7')
            ))
          )
        ),
        // Content: live preview
        h('div', { className: 'flex-1 overflow-hidden px-5 pb-5', style: { paddingTop: 8 } },
          h('div', { className: 'preview-wrap' },
            h('div', { className: 'preview-bar' },
              h('span', { className: 'preview-dot', style: { background: '#ef4444' } }),
              h('span', { className: 'preview-dot', style: { background: '#eab308' } }),
              h('span', { className: 'preview-dot', style: { background: '#22c55e' } }),
              h('span', { className: 'preview-url' }, window.location.origin + siteUrl),
              h('button', { onClick: () => { const url = window.location.origin + siteUrl; copyToClipboard(url, showToast); }, className: 'btn-ghost px-2.5 py-1 rounded-lg text-xs', title: 'Copy URL' }, '\u2398 Copy'),
              h('button', { onClick: () => setPreviewKey(k => k + 1), className: 'btn-ghost px-2.5 py-1 rounded-lg text-xs' }, '\u21BB Refresh')
            ),
            h('iframe', { key: previewKey, src: siteUrl, className: 'site-preview', title: 'Site preview' })
          )
        )
      ),

      // Right: chat
      h('div', { className: 'flex flex-col', style: { width: 400, minWidth: 340, borderLeft: '1px solid var(--border)', background: 'var(--bg)' } },
        h('div', { className: 'px-5 py-3.5', style: { borderBottom: '1px solid var(--border)' } },
          h('div', { className: 'flex items-center gap-2' },
            h('span', { style: { fontSize: 18 } }, '\u{2728}'),
            h('div', null,
              h('h3', { className: 'text-sm font-bold text-white' }, 'AI Builder'),
              h('p', { className: 'text-xs', style: { color: 'var(--text-mute)' } }, 'Describe your website, AI does the rest')
            )
          )
        ),
        h('div', { className: 'flex-1 overflow-y-auto px-4 py-4', style: { display: 'flex', flexDirection: 'column', gap: 12 } },
          messages.length === 0 && h('div', { className: 'empty-state', style: { padding: '40px 12px' } },
            h('div', { className: 'icon' }, '\u{1F4AC}'),
            h('p', { className: 'text-sm mb-1', style: { color: 'var(--text-dim)' } }, 'Start building'),
            h('p', { className: 'text-xs' }, 'E.g. "Maak een moderne landing page over honden met de foto\u2019s"')
          ),
          messages.map((msg, i) => msg.role === 'user'
            ? h(UserMessage, { key: msg.id || i, msg, showToast })
            : h(AiMessage, { key: msg.id || i, msg, showToast })
          ),
          generating && h('div', { className: 'chat-bubble chat-bubble-ai' },
            h('div', { className: 'flex items-center gap-3' },
              h('div', { className: 'spinner spinner-sm' }),
              h('span', { className: 'text-sm', style: { color: 'var(--text-dim)' } }, 'AI is building your website. This can take up to 5 to 10 minutes.')
            ),
            h('p', { className: 'text-xs mt-2', style: { color: 'var(--text-mute)' } },
              'You can safely leave this page \u2014 the generation continues in the background.' +
              (notifyEmail ? ' You will receive an email when it is done.' : '')
            ),
            jobId && h('button', {
              onClick: handleCancel,
              className: 'btn-ghost px-3 py-1.5 rounded-lg text-xs mt-2',
              style: { color: 'var(--danger)', borderColor: 'rgba(239,68,68,0.4)' }
            }, '\u2716 Cancel generation')
          ),
          h('div', { ref: chatEndRef })
        ),
        h('div', { className: 'p-4', style: { borderTop: '1px solid var(--border)' } },
          (() => {
            const isFirstPrompt = (backups.length === 0 && messages.filter(m => m.role === 'user').length === 0);
            const firstGenerationComplete = backups.length > 0;
            if (isFirstPrompt) {
              return h('div', { className: 'flex justify-start mb-2' },
                h('button', {
                  type: 'button',
                  onClick: () => setShowSuggestions(true),
                  className: 'btn-ghost px-3 py-1.5 rounded-lg text-xs font-semibold'
                }, 'Suggestions')
              );
            }
            if (firstGenerationComplete) {
              return h('div', { className: 'flex justify-start mb-2' },
                h('button', {
                  type: 'button',
                  onClick: () => setShowSuggestions(true),
                  className: 'btn-ghost px-3 py-1.5 rounded-lg text-xs font-semibold'
                }, 'Suggest improvement')
              );
            }
            return null;
          })(),
          h('form', { onSubmit: handleGenerate, className: 'flex gap-2 items-end' },
            h('textarea', {
              ref: promptRef,
              value: prompt,
              onChange: e => setPrompt(e.target.value),
              onKeyDown: handlePromptKeyDown,
              placeholder: 'Describe your website\u2026 (Enter = send, Shift+Enter = new line)',
              className: 'flex-1 text-sm resize-none prompt-input',
              rows: 4,
              disabled: generating
            }),
            generating
              ? h('button', {
                  type: 'button',
                  onClick: handleCancel,
                  disabled: !jobId,
                  title: jobId ? 'Cancel generation' : 'Cancelling is not available for synchronous generations',
                  className: 'px-4 py-2 rounded-xl text-sm font-semibold cancel-btn disabled:opacity-40'
                }, '\u2716')
              : h('button', {
                  type: 'submit',
                  disabled: !prompt.trim(),
                  className: 'gradient-btn px-4 py-2 text-white rounded-xl text-sm font-semibold disabled:opacity-40'
                }, '\u27A4')
          ),
          h('label', { className: 'flex items-center gap-2 mt-2.5 cursor-pointer', style: { userSelect: 'none' } },
            h('input', {
              type: 'checkbox',
              checked: notifyEmail,
              onChange: e => setNotifyEmail(e.target.checked),
              disabled: generating,
              style: { width: 14, height: 14, accentColor: 'var(--accent)' }
            }),
            h('span', { className: 'text-xs', style: { color: 'var(--text-dim)' } }, 'Email me when the generation is complete')
          )
        )
      )
    ),

    // Full-width backup preview modal
    previewBackup && h(BackupPreviewModal, {
      projectId,
      backupName: previewBackup,
      restoring,
      onClose: () => setPreviewBackup(null),
      onRestore: () => handleRestore(previewBackup)
    }),

    // Suggestions modal (bootstrap style, small)
    showSuggestions && h(SuggestionsModal, {
      isFirst: (backups.length === 0 && messages.filter(m => m.role === 'user').length === 0),
      onClose: () => setShowSuggestions(false),
      onSelect: (text) => {
        setPrompt(prev => {
          const trimmed = (prev || '').trim();
          return trimmed ? trimmed + ' ' + text : text;
        });
        setShowSuggestions(false);
        setTimeout(() => { if (promptRef.current) promptRef.current.focus(); }, 0);
      }
    })
  );
}

ReactDOM.createRoot(document.getElementById('root')).render(h(App));
