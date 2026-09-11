'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { createClient } from '@/lib/supabase/client';
import { Lock, Mail } from 'lucide-react';

export default function LoginPage() {
  const router = useRouter();
  const supabase = createClient();
  const [mode, setMode] = useState<'signin' | 'signup'>('signin');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    setError('');
    setNotice('');
    setLoading(true);
    try {
      if (mode === 'signin') {
        const { error } = await supabase.auth.signInWithPassword({ email, password });
        if (error) throw error;
        router.push('/');
        router.refresh();
      } else {
        const { data, error } = await supabase.auth.signUp({ email, password });
        if (error) throw error;
        if (data.session) {
          router.push('/');
          router.refresh();
        } else {
          setNotice('Conta criada. Verifique o email para confirmar antes de entrar.');
          setMode('signin');
        }
      }
    } catch (err: any) {
      setError(err?.message || 'Não foi possível concluir. Tente novamente.');
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="min-h-screen flex items-center justify-center bg-stone-100 p-4">
      <div className="w-full max-w-sm bg-white border border-stone-200 rounded p-6">
        <h1 className="font-serif-display text-2xl text-stone-800 mb-1">Gestão Imobiliária</h1>
        <p className="text-sm text-stone-500 mb-6">
          {mode === 'signin' ? 'Entre com a sua conta.' : 'Crie a sua conta de acesso.'}
        </p>

        <form onSubmit={handleSubmit}>
          <label className="block mb-3">
            <span className="block text-xs text-stone-500 mb-1">Email</span>
            <div className="flex items-center gap-2 border border-stone-300 rounded px-3 py-2">
              <Mail size={15} className="text-stone-400" />
              <input
                type="email" required value={email} onChange={(e) => setEmail(e.target.value)}
                className="w-full text-sm outline-none" placeholder="voce@exemplo.com"
              />
            </div>
          </label>
          <label className="block mb-4">
            <span className="block text-xs text-stone-500 mb-1">Password</span>
            <div className="flex items-center gap-2 border border-stone-300 rounded px-3 py-2">
              <Lock size={15} className="text-stone-400" />
              <input
                type="password" required minLength={6} value={password} onChange={(e) => setPassword(e.target.value)}
                className="w-full text-sm outline-none" placeholder="••••••••"
              />
            </div>
          </label>

          {error && <p className="text-sm text-red-700 mb-3">{error}</p>}
          {notice && <p className="text-sm text-emerald-700 mb-3">{notice}</p>}

          <button
            type="submit" disabled={loading}
            className="w-full rounded bg-emerald-800 text-white text-sm py-2.5 hover:bg-emerald-900 disabled:opacity-50"
          >
            {loading ? 'A processar…' : mode === 'signin' ? 'Entrar' : 'Criar conta'}
          </button>
        </form>

        <button
          onClick={() => { setMode(mode === 'signin' ? 'signup' : 'signin'); setError(''); setNotice(''); }}
          className="w-full text-center text-xs text-stone-500 hover:text-stone-700 mt-4"
        >
          {mode === 'signin' ? 'Primeira vez? Criar conta' : 'Já tem conta? Entrar'}
        </button>
      </div>
    </div>
  );
}
