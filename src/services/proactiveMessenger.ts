import {
  CloudAdapter,
  TurnContext,
  ConversationReference,
  Activity,
  CardFactory,
  MessageFactory,
} from 'botbuilder';
import { DataService } from './dataService';
import { NotificationService, Notification } from './notificationService';
import { buildStepListCard } from '../cards/adaptiveCards';

/**
 * Proactive messaging service for sending notifications to Teams users.
 * Maintains conversation references and sends messages outside of user-initiated turns.
 */
export class ProactiveMessenger {
  private adapter: CloudAdapter;
  private appId: string;
  private dataService: DataService;
  private notificationService: NotificationService;

  // Store conversation references for proactive messaging
  // Key: user email or Teams ID, Value: conversation reference
  private conversationRefs: Map<string, Partial<ConversationReference>> = new Map();

  constructor(
    adapter: CloudAdapter,
    appId: string,
    dataService: DataService,
    notificationService: NotificationService
  ) {
    this.adapter = adapter;
    this.appId = appId;
    this.dataService = dataService;
    this.notificationService = notificationService;
  }

  /** Store a conversation reference when a user messages the bot */
  saveConversationReference(activity: Activity): void {
    const ref = TurnContext.getConversationReference(activity);
    const userId = activity.from?.aadObjectId || activity.from?.id || '';
    if (userId) {
      this.conversationRefs.set(userId, ref);
      // Also store by name for lookup
      if (activity.from?.name) {
        this.conversationRefs.set(activity.from.name.toLowerCase(), ref);
      }
    }
  }

  /** Send a proactive message to a user by their stored conversation reference */
  async sendProactiveMessage(userId: string, message: string): Promise<boolean> {
    const ref = this.conversationRefs.get(userId) || this.conversationRefs.get(userId.toLowerCase());
    if (!ref) {
      console.warn(`[ProactiveMessenger] No conversation reference for ${userId}`);
      return false;
    }

    try {
      await this.adapter.continueConversationAsync(
        this.appId,
        ref as ConversationReference,
        async (context: TurnContext) => {
          await context.sendActivity(message);
        }
      );
      return true;
    } catch (err: any) {
      console.error(`[ProactiveMessenger] Failed to send to ${userId}: ${err.message}`);
      return false;
    }
  }

  /** Send an Adaptive Card proactively */
  async sendProactiveCard(userId: string, card: any): Promise<boolean> {
    const ref = this.conversationRefs.get(userId) || this.conversationRefs.get(userId.toLowerCase());
    if (!ref) return false;

    try {
      await this.adapter.continueConversationAsync(
        this.appId,
        ref as ConversationReference,
        async (context: TurnContext) => {
          await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
        }
      );
      return true;
    } catch (err: any) {
      console.error(`[ProactiveMessenger] Card send failed for ${userId}: ${err.message}`);
      return false;
    }
  }

  /** Process and deliver all pending notifications */
  async deliverNotifications(): Promise<{ sent: number; failed: number }> {
    const notifications = await this.notificationService.generateNotifications();
    let sent = 0;
    let failed = 0;

    for (const notification of notifications) {
      const recipientKey = notification.recipientTeamsId || notification.recipientEmail;
      const success = await this.sendProactiveMessage(
        recipientKey,
        this.formatNotification(notification)
      );
      if (success) sent++;
      else failed++;
    }

    console.log(`[ProactiveMessenger] Delivered ${sent}/${sent + failed} notifications`);
    return { sent, failed };
  }

  /** Process and deliver weekly digest */
  async deliverWeeklyDigest(): Promise<{ sent: number; failed: number }> {
    const digests = await this.notificationService.generateWeeklyDigest();
    let sent = 0;
    let failed = 0;

    for (const digest of digests) {
      const recipientKey = digest.recipientTeamsId || digest.recipientEmail;
      const success = await this.sendProactiveMessage(recipientKey, digest.message);
      if (success) sent++;
      else failed++;
    }

    console.log(`[ProactiveMessenger] Weekly digest: ${sent}/${sent + failed} delivered`);
    return { sent, failed };
  }

  /** Notify relevant users when a predecessor step completes */
  async deliverPredecessorNotifications(completedStepId: string, track: 'Corp' | 'Fed'): Promise<number> {
    const notifications = await this.notificationService.notifyPredecessorComplete(completedStepId, track);
    let sent = 0;

    for (const n of notifications) {
      const recipientKey = n.recipientTeamsId || n.recipientEmail;
      const success = await this.sendProactiveMessage(recipientKey, this.formatNotification(n));
      if (success) sent++;
    }

    return sent;
  }

  /** Get count of stored conversation references */
  getRegisteredUserCount(): number {
    return this.conversationRefs.size;
  }

  /** Check if we have a conversation reference for a user */
  hasConversationRef(userId: string): boolean {
    return this.conversationRefs.has(userId) || this.conversationRefs.has(userId.toLowerCase());
  }

  /** Format a notification into a readable Teams message */
  private formatNotification(notification: Notification): string {
    const priorityIcon = notification.priority === 'high' ? '🔴' :
      notification.priority === 'medium' ? '🟡' : '🟢';

    let msg = `${priorityIcon} **${notification.title}**\n\n`;
    msg += notification.message;

    if (notification.steps.length > 0) {
      msg += '\n\n---\n';
      for (const step of notification.steps.slice(0, 5)) {
        msg += `• **${step.id}** ${step.description} → ${step.corpStatus}\n`;
      }
    }

    return msg;
  }
}
