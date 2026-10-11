import { PlusIcon, TrashIcon } from '@phosphor-icons/react'
import { useState } from 'react'
import type { ProviderKind, ProviderSpec } from '../../shared/api/ai'
import { Button, IconButton, InlineAlert, Select, TextField, Working } from '../components/ui'
import { call, useAction, useApi } from '../lib/api'
import { bridge } from '../lib/bridge'
import { toast } from '../lib/toast'
import { t } from '../strings/en'

/** Adds a model provider, checks it answers, and makes it the model for every role. */
export function ProviderForm({ onDone }: { onDone: () => void }) {
  const catalog = useApi('ai.catalog', undefined)
  const local = useApi('ai.detectLocal', undefined)
  const [kind, setKind] = useState<ProviderKind>('anthropic')
  const [fields, setFields] = useState<Record<string, string>>({})
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const spec: ProviderSpec | undefined = catalog.data?.find((c) => c.kind === kind)
  const save = async () => {
    setBusy(true)
    setError(null)
    try {
      const p = await call('ai.saveProvider', { kind, fields })
      const withModels = p.models.length ? p : await call('ai.refreshModels', { id: p.id }).catch(() => p)
      await call('ai.autoAssign', { providerId: withModels.id, onlyEmpty: false })
      onDone()
    } catch (e) {
      setError((e as Error).message)
    } finally {
      setBusy(false)
    }
  }
  const found = [local.data?.ollama?.length ? `Ollama: ${local.data.ollama.slice(0, 3).join(', ')}` : '', local.data?.lmstudio?.length ? `LM Studio: ${local.data.lmstudio.slice(0, 3).join(', ')}` : ''].filter(Boolean)
  return (
    <div className="stack">
      <Select label={t.settings.models.kind} value={kind} onChange={(k) => { setKind(k); setFields({}) }} options={(catalog.data ?? []).filter((c) => c.kind !== 'mock').map((c) => ({ id: c.kind, label: c.label }))} />
      {found.length > 0 && <p className="meta">{found.join('. ')}</p>}
      {spec?.fields.map((f) => (
        <TextField
          key={f.key}
          label={f.label}
          help={f.help || undefined}
          placeholder={f.placeholder || undefined}
          type={f.secret ? 'password' : 'text'}
          multiline={f.multiline}
          isRequired={f.required}
          value={fields[f.key] ?? ''}
          onChange={(v) => setFields({ ...fields, [f.key]: v })}
        />
      ))}
      <div className="row">
        <Button variant="primary" isDisabled={busy || !spec || spec.fields.some((f) => f.required && !fields[f.key]?.trim())} onPress={() => void save()}>
          {t.settings.models.add}
        </Button>
        {spec?.keyUrl && (
          <Button variant="quiet" onPress={() => void bridge.host.openExternal(spec.keyUrl!)}>
            {t.settings.integrations.getKey}
          </Button>
        )}
      </div>
      {busy && <Working text={t.settings.models.testing} />}
      {error && <InlineAlert tone="danger">{error}</InlineAlert>}
    </div>
  )
}

/**
 * The onboarding model step: shows every saved provider (it may have failed to answer, so each can
 * be tested, refreshed or removed), lets another be added, and only continues once a role has a model.
 */
export function ProviderSetup({ onDone }: { onDone: () => void }) {
  const providers = useApi('ai.providers', undefined)
  const roles = useApi('ai.roles', undefined)
  const del = useAction('ai.deleteProvider')
  const test = useAction('ai.test')
  const refresh = useAction('ai.refreshModels')
  const [adding, setAdding] = useState(false)
  const list = providers.data ?? []
  const ready = (roles.data ?? []).some((r) => r.modelId !== null)
  if (list.length === 0) return <ProviderForm onDone={onDone} />
  return (
    <div className="stack">
      <ul className="panel">
        {list.map((p) => (
          <li key={p.id} className="worklist-item" style={{ gridTemplateColumns: 'minmax(0,1fr) auto' }}>
            <div className="stack" style={{ gap: 2 }}>
              <span className="what">{p.label}</span>
              <span className="why">{p.models.length ? t.settings.models.modelCount(p.models.length) : t.settings.models.noModels}</span>
            </div>
            <div className="row">
              {p.models[0] ? (
                <Button size="sm" isDisabled={test.isPending} onPress={() => test.mutate({ providerId: p.id, modelId: p.models[0]! }, { onSuccess: (r) => (r.ok ? toast.info(t.settings.models.testOk(r.latencyMs, p.models[0]!)) : toast.error(r.error ?? '')) })}>
                  {t.settings.models.test}
                </Button>
              ) : (
                <Button size="sm" isDisabled={refresh.isPending} onPress={() => refresh.mutate({ id: p.id })}>
                  {t.settings.models.refresh}
                </Button>
              )}
              <IconButton size="sm" label={t.settings.models.remove} icon={<TrashIcon aria-hidden />} onPress={() => del.mutate({ id: p.id })} />
            </div>
          </li>
        ))}
      </ul>
      {adding ? (
        <ProviderForm onDone={() => setAdding(false)} />
      ) : (
        <div className="row">
          <Button onPress={() => setAdding(true)}>
            <PlusIcon aria-hidden />
            {t.settings.models.add}
          </Button>
          <span className="spacer" />
          <Button variant="primary" isDisabled={!ready} onPress={onDone}>
            {t.onboarding.next}
          </Button>
        </div>
      )}
    </div>
  )
}

/** Connects a mailbox: app password for most providers, a browser sign-in for Microsoft. */
export function MailForm({ onDone }: { onDone: () => void }) {
  const presets = useApi('mail.presets', undefined)
  const [preset, setPreset] = useState('gmail')
  const [address, setAddress] = useState('')
  const [password, setPassword] = useState('')
  const [name, setName] = useState('')
  const [clientId, setClientId] = useState('')
  const [imap, setImap] = useState({ host: '', port: '993' })
  const [smtp, setSmtp] = useState({ host: '', port: '465' })
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const p = presets.data?.find((x) => x.id === preset)
  const run = async () => {
    setBusy(true)
    setError(null)
    try {
      if (p?.auth === 'oauth') await call('mail.connectOAuth', { provider: 'microsoft', clientId, address: address || undefined, displayName: name || undefined })
      else
        await call('mail.connect', {
          preset,
          address,
          password,
          displayName: name || undefined,
          ...(p?.needsServers ? { imap: { host: imap.host, port: Number(imap.port), secure: Number(imap.port) === 993 }, smtp: { host: smtp.host, port: Number(smtp.port), secure: Number(smtp.port) === 465 } } : {}),
        })
      onDone()
    } catch (e) {
      setError((e as Error).message)
    } finally {
      setBusy(false)
    }
  }
  return (
    <div className="stack">
      <Select label={t.settings.email.provider} value={preset} onChange={setPreset} options={(presets.data ?? []).map((x) => ({ id: x.id, label: x.label }))} />
      <div className="grid-2">
        <TextField label={t.settings.email.address} type="email" value={address} onChange={setAddress} />
        <TextField label={t.settings.email.displayName} value={name} onChange={setName} />
      </div>
      {p?.auth === 'oauth' ? (
        <TextField label={t.settings.email.clientId} help={t.settings.email.clientIdHelp} value={clientId} onChange={setClientId} />
      ) : (
        <TextField label={t.settings.email.password} help={t.settings.email.passwordHelp} type="password" value={password} onChange={setPassword} />
      )}
      {p?.needsServers && (
        <div className="grid-2">
          <div className="row">
            <TextField label={t.settings.email.imapHost} value={imap.host} onChange={(v) => setImap({ ...imap, host: v })} className="spacer" />
            <TextField label={t.settings.email.port} value={imap.port} onChange={(v) => setImap({ ...imap, port: v })} />
          </div>
          <div className="row">
            <TextField label={t.settings.email.smtpHost} value={smtp.host} onChange={(v) => setSmtp({ ...smtp, host: v })} className="spacer" />
            <TextField label={t.settings.email.port} value={smtp.port} onChange={(v) => setSmtp({ ...smtp, port: v })} />
          </div>
        </div>
      )}
      <div className="row">
        <Button variant="primary" isDisabled={busy || !address.includes('@') || (p?.auth === 'oauth' ? clientId.length < 8 : !password)} onPress={() => void run()}>
          {p?.auth === 'oauth' ? t.settings.email.signIn : t.settings.email.add}
        </Button>
        {p?.helpUrl && (
          <Button variant="quiet" onPress={() => void bridge.host.openExternal(p.helpUrl!)}>
            {p.auth === 'oauth' ? t.common.showDetails : t.settings.email.makePassword}
          </Button>
        )}
      </div>
      {busy && <Working text={t.settings.email.connecting} />}
      {error && <InlineAlert tone="danger">{error}</InlineAlert>}
    </div>
  )
}
