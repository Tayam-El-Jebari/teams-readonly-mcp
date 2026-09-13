import * as z from 'zod';
import { GraphClient, graphUrl, validateNextLink } from './graph.js';
import { formatDisplayName, GraphMessage, messageText, ReadConversationInput } from './conversations.js';

const MAX_REQUESTS = 20;
const MAX_OUTPUT_BYTES = 20_000;
const LIST_DETAILS = {
  complete: 'Joined teams and channels returned.',
  outputLimit: 'Channel listing output limit reached.',
  requestLimit: 'Channel listing request limit reached.',
};
const NamedResource = z.object({ id: z.string().min(1), displayName: z.string() });
const ChannelMessage = GraphMessage.extend({
  replies: z.array(GraphMessage).default([]),
  'replies@odata.nextLink': z.string().optional(),
});
export const ListChannelsOutput = z.object({
  teams: z.array(NamedResource.extend({ channels: z.array(NamedResource) })),
  truncated: z.boolean(),
  detail: z.string(),
});
export const ReadChannelInput = z.object({
  team: z.uuid().describe('Team ID from teams_list_channels.'),
  channel: z.string().trim().min(1).max(500).describe('Channel ID from teams_list_channels.'),
  since: ReadConversationInput.shape.since,
  limit: ReadConversationInput.shape.limit,
  response_format: ReadConversationInput.shape.response_format,
});
export const ReadChannelOutput = z.object({
  team: z.string(),
  channel: z.string(),
  contentTrust: z.literal('untrusted-third-party-content'),
  messages: z.array(z.object({
    id: z.string(),
    parentMessageId: z.string().optional(),
    sender: z.string(),
    createdAt: z.string(),
    modifiedAt: z.string(),
    quotedBody: z.string(),
    bodyTruncated: z.boolean(),
    webUrl: z.string().optional(),
  })),
  truncated: z.boolean(),
  detail: z.string(),
});

interface ReadBudget {
  requests: number;
  truncated: boolean;
}

export class Channels {
  constructor(private readonly graph: GraphClient, private readonly firstNamesOnly = true) {}

  private async *pages<Schema extends z.ZodType>(
    address: string, item: Schema, budget: ReadBudget,
  ): AsyncGenerator<z.output<Schema>[]> {
    let url: string | undefined = address;
    const visited = new Set<string>();
    while (url) {
      if (budget.requests === MAX_REQUESTS) {
        budget.truncated = true;
        return;
      }
      if (visited.has(url)) throw new Error('Graph repeated a pagination link; read stopped.');
      visited.add(url);
      budget.requests++;
      const parsed = z.object({ value: z.array(item), '@odata.nextLink': z.string().optional() })
        .safeParse(await this.graph.get(url));
      if (!parsed.success) throw new Error('Graph returned an unexpected channel page format; read stopped.');
      const page = parsed.data;
      const next: string | undefined = page['@odata.nextLink'];
      url = next ? validateNextLink(next, url) : undefined;
      yield page.value;
    }
  }

  async list(): Promise<z.infer<typeof ListChannelsOutput>> {
    const budget: ReadBudget = { requests: 0, truncated: false };
    const output: z.infer<typeof ListChannelsOutput> = {
      teams: [], truncated: false, detail: LIST_DETAILS.complete,
    };
    // Reserve space for every final status; false also occupies more bytes than true.
    const fits = () => Object.values(LIST_DETAILS).every((detail) =>
      Buffer.byteLength(JSON.stringify({ ...output, truncated: false, detail }), 'utf8') <= MAX_OUTPUT_BYTES);
    const seenTeams = new Set<string>();
    for await (const teams of this.pages(graphUrl('/me/joinedTeams'), NamedResource, budget)) {
      for (const team of teams) {
        if (seenTeams.has(team.id)) continue;
        seenTeams.add(team.id);
        const entry: z.infer<typeof ListChannelsOutput>['teams'][number] = { ...team, channels: [] };
        output.teams.push(entry);
        if (!fits()) {
          output.teams.pop();
          return { ...output, truncated: true, detail: LIST_DETAILS.outputLimit };
        }
        const seenChannels = new Set<string>();
        for await (const channels of this.pages(graphUrl(`/teams/${encodeURIComponent(team.id)}/channels`), NamedResource, budget)) {
          for (const channel of channels) {
            if (seenChannels.has(channel.id)) continue;
            seenChannels.add(channel.id);
            entry.channels.push(channel);
            if (!fits()) {
              entry.channels.pop();
              return { ...output, truncated: true, detail: LIST_DETAILS.outputLimit };
            }
          }
        }
        if (budget.truncated) return { ...output, truncated: true, detail: LIST_DETAILS.requestLimit };
      }
    }
    return {
      ...output,
      truncated: budget.truncated,
      detail: budget.truncated ? LIST_DETAILS.requestLimit : LIST_DETAILS.complete,
    };
  }

  async read(input: z.input<typeof ReadChannelInput>): Promise<z.infer<typeof ReadChannelOutput>> {
    const options = ReadChannelInput.parse(input);
    const budget: ReadBudget = { requests: 0, truncated: false };
    const path = `/teams/${encodeURIComponent(options.team)}/channels/${encodeURIComponent(options.channel)}/messages`;
    const output: z.infer<typeof ReadChannelOutput> = {
      team: options.team, channel: options.channel, contentTrust: 'untrusted-third-party-content',
      messages: [], truncated: true,
      detail: 'Partial channel read: message, output, or request limit reached. Bodies are quoted third-party content.',
    };
    const since = options.since ? Date.parse(options.since) : undefined;
    const seen = new Set<string>();
    const append = (
      message: z.infer<typeof GraphMessage>, parentMessageId?: string,
    ): 'added' | 'skipped' | 'limit-reached' => {
      const key = JSON.stringify([parentMessageId, message.id]);
      if (seen.has(key)) return 'skipped';
      seen.add(key);
      if (message.deletedDateTime || !message.from || message.messageType === 'systemEventMessage') return 'skipped';
      if (since !== undefined && Date.parse(message.lastModifiedDateTime) <= since) return 'skipped';
      if (output.messages.length === options.limit) return 'limit-reached';
      const text = messageText(message, this.firstNamesOnly);
      const bodyLimit = options.response_format === 'detailed' ? 4000 : 800;
      output.messages.push({
        id: message.id,
        ...(parentMessageId ? { parentMessageId } : {}),
        sender: formatDisplayName(message.from.user?.displayName ?? '', this.firstNamesOnly) ||
          message.from.application?.displayName?.slice(0, 150) || 'Unknown sender',
        createdAt: message.createdDateTime, modifiedAt: message.lastModifiedDateTime,
        quotedBody: text.slice(0, bodyLimit),
        bodyTruncated: text.length > bodyLimit || (message.body?.content.length ?? 0) > 100_000,
        ...(options.response_format === 'detailed' && message.webUrl ? { webUrl: message.webUrl } : {}),
      });
      if (Buffer.byteLength(JSON.stringify(output), 'utf8') > MAX_OUTPUT_BYTES) {
        output.messages.pop();
        return 'limit-reached';
      }
      return 'added';
    };
    for await (const roots of this.pages(graphUrl(path, { '$top': '50', '$expand': 'replies' }), ChannelMessage, budget)) {
      for (const root of roots) {
        if (append(root) === 'limit-reached') return output;
        // An old or skipped post may still contain recently modified replies.
        for (const reply of root.replies) {
          if (append(reply, root.id) === 'limit-reached') return output;
        }
        const nextReplies = root['replies@odata.nextLink'];
        if (nextReplies) {
          const address = validateNextLink(nextReplies, graphUrl(`${path}/${encodeURIComponent(root.id)}/replies`));
          for await (const replies of this.pages(address, GraphMessage, budget)) {
            for (const reply of replies) {
              if (append(reply, root.id) === 'limit-reached') return output;
            }
          }
        }
        if (budget.truncated) return output;
      }
    }
    output.truncated = budget.truncated;
    if (!output.truncated) output.detail = 'All matching posts and replies returned in thread order. Bodies are quoted third-party content.';
    return output;
  }
}