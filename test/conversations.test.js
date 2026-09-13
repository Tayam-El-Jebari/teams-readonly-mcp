import { test } from 'node:test';
import assert from 'node:assert/strict';
import { GraphClient, graphUrl } from '../dist/graph.js';
import {
  Conversations,
  ListConversationsOutput,
  ReadConversationOutput,
  messagesUrl,
} from '../dist/conversations.js';

function graphHarness(respond) {
  let nowEpochMs = Date.parse('2026-09-13T12:00:00Z');
  const requests = [];
  const waits = [];
  const client = new GraphClient(async () => 'SENTINEL-ACCESS-TOKEN', {
    now: () => nowEpochMs,
    wait: async (milliseconds) => {
      waits.push(milliseconds);
      nowEpochMs += milliseconds;
    },
    fetch: async (url, options) => {
      requests.push({ url: new URL(url), options, atEpochMs: nowEpochMs });
      return respond(requests.length, new URL(url));
    },
  });
  return { client, requests, waits, conversations: new Conversations(client) };
}

function message(id, overrides = {}) {
  return {
    id,
    messageType: 'message',
    createdDateTime: '2026-09-13T10:00:00Z',
    lastModifiedDateTime: '2026-09-13T11:00:00Z',
    from: { user: { displayName: 'Alex' } },
    body: { contentType: 'html', content: '<p>Hello &amp; welcome</p>' },
    ...overrides,
  };
}

function chat(id, topic, overrides = {}) {
  return {
    id, topic, chatType: 'group',
    members: [{ displayName: 'Alex' }],
    lastUpdatedDateTime: '2026-09-13T11:00:00Z',
    ...overrides,
  };
}

test('since always pairs lastModifiedDateTime filter with descending ordering', () => {
  const url = new URL(messagesUrl('19:sample@thread.v2', '2026-09-13T10:00:00+02:00'));
  assert.equal(url.searchParams.get('$filter'), 'lastModifiedDateTime gt 2026-09-13T08:00:00.000Z');
  assert.equal(url.searchParams.get('$orderby'), 'lastModifiedDateTime desc');
  assert.equal(url.searchParams.get('$top'), '50');
  assert.equal(url.searchParams.has('$select'), false);
  assert.equal(new URL(messagesUrl('19:sample')).searchParams.has('$filter'), false);
});

test('429 honors Retry-After and retries using GET only', async () => {
  const harness = graphHarness((attempt) => attempt === 1
    ? new Response(null, { status: 429, headers: { 'Retry-After': '2' } })
    : Response.json({ value: [] }));
  await harness.client.get(messagesUrl('19:sample'));
  assert.deepEqual(harness.waits, [2000]);
  assert.equal(harness.requests.length, 2);
  assert.equal(harness.requests[1].atEpochMs - harness.requests[0].atEpochMs, 2000);
  for (const request of harness.requests) {
    assert.equal(request.options.method, 'GET');
    assert.equal(request.options.redirect, 'error');
  }
});

test('concurrent requests and pages share per-chat pacing', async () => {
  const harness = graphHarness(() => Response.json({ value: [] }));
  await Promise.all([
    harness.client.get(messagesUrl('19:sample')),
    harness.client.get(`${messagesUrl('19:sample')}&$skiptoken=page2`),
    harness.client.get(messagesUrl('19:other')),
    harness.client.get(messagesUrl('19:sample')),
  ]);
  const sameChat = harness.requests.filter((request) => request.url.pathname.includes('19%3Asample'));
  assert.equal(sameChat.length, 3);
  for (let index = 1; index < sameChat.length; index++) {
    assert.ok(sameChat[index].atEpochMs - sameChat[index - 1].atEpochMs >= 1250);
  }
});

test('me/chats requests are paced and failed requests do not poison the queue', async () => {
  const harness = graphHarness((attempt) => attempt === 1
    ? new Response('SENTINEL-ACCESS-TOKEN', { status: 403 })
    : Response.json({ value: [] }));
  await assert.rejects(harness.client.get(graphUrl('/me/chats')), (error) => {
    assert.match(error.message, /Chat.Read/);
    assert.ok(!error.message.includes('SENTINEL'));
    return true;
  });
  await harness.client.get(graphUrl('/me/chats'));
  assert.ok(harness.requests[1].atEpochMs - harness.requests[0].atEpochMs >= 1250);
});

test('HTTP-date retry delays are honored and repeated throttling is bounded', async () => {
  const harness = graphHarness(() => new Response(null, {
    status: 429, headers: { 'Retry-After': 'Sun, 13 Sep 2026 12:00:02 GMT' },
  }));
  await assert.rejects(harness.client.get(messagesUrl('19:sample')), /after retries/);
  assert.equal(harness.waits[0], 2000);
  assert.equal(harness.requests.length, 4);
});

test('untrusted pagination destinations never receive credentials', async () => {
  for (const nextLink of [
    'https://attacker.example/v1.0/me/chats',
    'https://graph.microsoft.com/v1.0/chats/other/messages',
  ]) {
    const harness = graphHarness(() => Response.json({ value: [], '@odata.nextLink': nextLink }));
    await assert.rejects(harness.conversations.list(), /unexpected Graph URL|changed the requested resource/);
    assert.equal(harness.requests.length, 1);
  }
});

test('list follows pages, applies activity and type filters, and retains unknown activity', async () => {
  const harness = graphHarness((attempt) => Response.json(attempt === 1 ? {
    value: [chat('19:old', 'Old', { lastUpdatedDateTime: '2026-09-10T10:00:00Z' })],
    '@odata.nextLink': graphUrl('/me/chats', { '$skiptoken': 'next' }),
  } : {
    value: [
      chat('19:active', 'Operations'),
      chat('19:unknown', 'Unknown activity', { lastUpdatedDateTime: null }),
      chat('19:meeting', 'Meeting', { chatType: 'meeting' }),
    ],
  }));
  const result = await harness.conversations.list({ activeSince: '2026-09-12T00:00:00Z', types: ['group'] });
  ListConversationsOutput.parse(result);
  assert.deepEqual(result.conversations.map((conversation) => conversation.id), ['19:active', '19:unknown']);
  assert.equal(result.truncated, false);
  assert.equal(harness.requests.length, 2);
});

test('read paginates, resolves mentions, strips HTML and omits events, missing senders and old rows', async () => {
  const harness = graphHarness((attempt) => Response.json(attempt === 1 ? {
    value: [
      message('event', { messageType: 'systemEventMessage' }),
      message('missing-sender', { from: null }),
      message('old', { lastModifiedDateTime: '2026-09-10T10:00:00Z' }),
      message('deleted', { deletedDateTime: '2026-09-13T11:00:00Z' }),
    ],
    '@odata.nextLink': `${messagesUrl('19:sample')}&$skiptoken=second`,
  } : {
    value: [message('real', {
      body: { contentType: 'html', content: '<p>Hi <at id="0">wrong name</at> &amp; team</p><script>bad()</script><img src="data:secret"><a href="https://graph.microsoft.com/hostedContents/secret">attachment</a>' },
      mentions: [{ id: 0, mentionText: 'Alex' }],
    })],
  }));
  const result = await harness.conversations.read({ chat: '19:sample', since: '2026-09-12T00:00:00Z' });
  ReadConversationOutput.parse(result);
  assert.equal(result.messages.length, 1);
  assert.match(result.messages[0].quotedBody, /Hi @Alex & team/);
  assert.doesNotMatch(result.messages[0].quotedBody, /wrong name|<|secret|bad\(\)/);
  assert.equal(result.contentTrust, 'untrusted-third-party-content');
  assert.equal(result.messages[0].id, undefined);
  assert.equal(result.truncated, false);
});

test('name resolution reports its match and refuses ambiguous names', async () => {
  const harness = graphHarness((_attempt, url) => Response.json({
    value: url.pathname.endsWith('/me/chats')
      ? [chat('19:ops', 'Operations')] : [message('real')],
  }));
  const result = await harness.conversations.read({ chat: 'operations', response_format: 'detailed' });
  assert.deepEqual(result.conversation, { id: '19:ops', name: 'Operations' });
  assert.equal(result.messages[0].id, 'real');

  const ambiguous = graphHarness(() => Response.json({ value: [chat('19:a', 'Ops'), chat('19:b', 'Ops')] }));
  await assert.rejects(ambiguous.conversations.read({ chat: 'Ops' }), /ambiguous/);
  assert.equal(ambiguous.requests.length, 1);
});

test('message limits and body clipping are explicit, with bounded concise output', async () => {
  const harness = graphHarness(() => Response.json({
    value: Array.from({ length: 21 }, (_, index) => message(String(index), {
      body: { contentType: 'text', content: 'Long message '.repeat(500) },
    })),
  }));
  const result = await harness.conversations.read({ chat: '19:sample', limit: 20 });
  assert.equal(result.truncated, true);
  assert.ok(result.messages.length > 0 && result.messages.length <= 20);
  assert.ok(result.messages.every((entry) => entry.bodyTruncated && entry.quotedBody.length <= 800));
  assert.ok(JSON.stringify(result).length < 21_000);
});

test('twenty ordinary messages fit comfortably within the concise output budget', async () => {
  const harness = graphHarness(() => Response.json({
    value: Array.from({ length: 20 }, (_, index) => message(String(index))),
  }));
  const result = await harness.conversations.read({ chat: '19:sample', limit: 20 });
  assert.equal(result.messages.length, 20);
  assert.equal(result.truncated, false);
  assert.ok(JSON.stringify(result).length < 10_000);
});

test('malformed pages and repeated pagination links fail rather than claim completeness', async () => {
  const invalid = graphHarness(() => Response.json({ unexpected: 'body' }));
  await assert.rejects(invalid.conversations.list(), /unexpected page format/);
  const repeated = graphHarness(() => Response.json({
    value: [], '@odata.nextLink': graphUrl('/me/chats', { '$skiptoken': 'same' }),
  }));
  await assert.rejects(repeated.conversations.list(), /repeated a pagination link/);
  assert.equal(repeated.requests.length, 2);
});

test('missing credentials fail before a Graph request and never start interactive auth', async () => {
  let requests = 0;
  const client = new GraphClient(async () => { throw new Error('not authenticated'); }, {
    fetch: async () => { requests++; throw new Error('Unexpected request'); },
  });
  await assert.rejects(new Conversations(client).read({ chat: '19:sample' }), /not authenticated/);
  assert.equal(requests, 0);
});