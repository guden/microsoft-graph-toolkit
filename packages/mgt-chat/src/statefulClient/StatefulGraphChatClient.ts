import {
  MessageThreadProps,
  SendBoxProps,
  ChatMessage as ACSChatMessage,
  ErrorBarProps
} from '@azure/communication-react';
import { Chat, ChatMessage } from '@microsoft/microsoft-graph-types';
import { ActiveAccountChanged, IGraph, LoginChangedEvent, Providers, ProviderState } from '@microsoft/mgt-element';
import { produce } from 'immer';
import { v4 as uuid } from 'uuid';
import {
  deleteChatMessage,
  loadChat,
  loadChatThread,
  loadMoreChatMessages,
  MessageCollection,
  sendChatMessage,
  updateChatMessage
} from './graph.chat';
import { graphChatMessageToACSChatMessage } from './acs.chat';
import { GraphNotificationClient } from './GraphNotificationClient';
import { ThreadEventEmitter } from './ThreadEventEmitter';

type GraphChatClient = Pick<
  MessageThreadProps,
  | 'userId'
  | 'messages'
  | 'participantCount'
  | 'disableEditing'
  | 'onLoadPreviousChatMessages'
  | 'numberOfChatMessagesToReload'
  | 'onUpdateMessage'
  | 'onDeleteMessage'
> &
  Pick<SendBoxProps, 'onSendMessage'> &
  Pick<ErrorBarProps, 'activeErrorMessages'> & {
    status: 'initial' | 'subscribing to notifications' | 'loading messages' | 'ready' | 'error';
    chat?: Chat;
  };

type StatefulClient<T> = {
  /**
   * Get the current state of the client
   */
  getState(): T;
  /**
   * Register a callback to receive state updates
   *
   * @param handler Callback to receive state updates
   */
  onStateChange(handler: (state: T) => void): void;
  /**
   * Remove a callback from receiving state updates
   *
   * @param handler Callback to be unregistered
   */
  offStateChange(handler: (state: T) => void): void;
};

type CreatedOn = {
  createdOn: Date;
};

/**
 * Simple object comparator function for sorting by createdOn date
 *
 * @param {CreatedOn} a
 * @param {CreatedOn} b
 */
const MessageCreatedComparator = (a: CreatedOn, b: CreatedOn) => a.createdOn.getTime() - b.createdOn.getTime();

class StatefulGraphChatClient implements StatefulClient<GraphChatClient> {
  private _notificationClient: GraphNotificationClient;
  private _eventEmitter: ThreadEventEmitter;
  private _subscribers: ((state: GraphChatClient) => void)[] = [];
  private _messagesPerCall = 5;
  private _nextLink = '';
  private _chat?: Chat = undefined;
  private _userDisplayName = '';

  constructor() {
    this.updateUserInfo();
    Providers.globalProvider.onStateChanged(this.onLoginStateChanged);
    Providers.globalProvider.onActiveAccountChanged(this.onActiveAccountChanged);
    this._notificationClient = new GraphNotificationClient();
    this._eventEmitter = new ThreadEventEmitter();
    this.registerEventListeners();
  }

  /**
   * Register a callback to receive state updates
   *
   * @param {(state: GraphChatClient) => void} handler
   * @memberof StatefulGraphChatClient
   */
  public onStateChange(handler: (state: GraphChatClient) => void): void {
    if (!this._subscribers.includes(handler)) {
      this._subscribers.push(handler);
    }
  }

  /**
   * Unregister a callback from receiving state updates
   *
   * @param {(state: GraphChatClient) => void} handler
   * @memberof StatefulGraphChatClient
   */
  public offStateChange(handler: (state: GraphChatClient) => void): void {
    const index = this._subscribers.indexOf(handler);
    if (index !== -1) {
      this._subscribers = this._subscribers.splice(index, 1);
    }
  }

  /**
   * Calls each subscriber with the next state to be emitted
   *
   * @param recipe - a function which produces the next state to be emitted
   */
  private notifyStateChange(recipe: (draft: GraphChatClient) => void) {
    this._state = produce(this._state, recipe);
    this._subscribers.forEach(handler => handler(this._state));
  }

  /**
   * Return the current state of the chat client
   *
   * @return {{GraphChatClient}
   * @memberof StatefulGraphChatClient
   */
  public getState(): GraphChatClient {
    return this._state;
  }

  /**
   * Update the state of the client when the Login state changes
   *
   * @private
   * @param {LoginChangedEvent} e The event that triggered the change
   * @memberof StatefulGraphChatClient
   */
  private onLoginStateChanged = (e: LoginChangedEvent) => {
    switch (Providers.globalProvider.state) {
      case ProviderState.SignedIn:
        // update userId and displayName
        this.updateUserInfo();
        // load messages?
        // configure subscriptions
        // emit new state;
        if (this._chatId) {
          void this.updateFollowedChat();
        }
        return;
      case ProviderState.SignedOut:
        // clear userId
        // clear subscriptions
        // clear messages
        // emit new state
        return;
      case ProviderState.Loading:
      default:
        // do nothing for now
        return;
    }
  };

  private onActiveAccountChanged = (e: ActiveAccountChanged) => {
    this.updateUserInfo();
  };

  private updateUserInfo() {
    this.updateCurrentUserId();
    this.updateCurrentUserName();
  }

  private updateCurrentUserName() {
    this._userDisplayName = Providers.globalProvider.getActiveAccount?.().name || '';
  }

  private updateCurrentUserId() {
    this.userId = Providers.globalProvider.getActiveAccount?.().id.split('.')[0] || '';
  }

  private _userId = '';
  private set userId(userId: string) {
    if (this._userId === userId) {
      return;
    }
    this._userId = userId;
    this.notifyStateChange((draft: GraphChatClient) => {
      draft.userId = userId;
    });
  }

  private _chatId = '';

  public set chatId(value: string) {
    // take no action if the chatId is the same
    if (this._chatId === value) {
      return;
    }
    this._chatId = value;
    void this.updateFollowedChat();
  }

  /**
   * A helper to co-ordinate the loading of a chat and its messages, and the subscription to notifications for that chat
   *
   * @private
   * @memberof StatefulGraphChatClient
   */
  private async updateFollowedChat() {
    // Subscribe to notifications for messages
    this.notifyStateChange((draft: GraphChatClient) => {
      draft.status = 'subscribing to notifications';
    });
    // subscribing to notifications will trigger the chatMessageNotificationsSubscribed event
    // this client will then load the chat and messages when that event listener is called
    await this._notificationClient.subscribeToChatNotifications(this._userId, this._chatId, this._eventEmitter);
  }

  private async loadChatData() {
    this.notifyStateChange((draft: GraphChatClient) => {
      draft.status = 'loading messages';
    });
    this._chat = await loadChat(this.graph, this._chatId);
    const messages: MessageCollection = await loadChatThread(this.graph, this._chatId, this._messagesPerCall);
    this._nextLink = messages.nextLink;
    // Allow messages to be loaded via the loadMoreMessages callback
    this.notifyStateChange((draft: GraphChatClient) => {
      draft.participantCount = this._chat?.members?.length || 0;
      draft.messages = messages.value
        // trying to filter out system messages on the graph request causes a 400
        // delted messages are returned as messages with no content, which we can't filter on the graph request
        // so we filter them out here
        .filter(m => m.messageType === 'message' && m.body?.content)
        .map(m => graphChatMessageToACSChatMessage(m, this._userId));
      draft.onLoadPreviousChatMessages = this._nextLink ? this.loadMoreMessages : undefined;
      draft.status = this._nextLink ? 'loading messages' : 'ready';
      draft.chat = this._chat;
    });
  }

  /**
   * Async callback to load more messages
   *
   * @returns true if there are no more messages to load
   */
  private loadMoreMessages = async () => {
    if (!this._nextLink) {
      return true;
    }
    const result: MessageCollection = await loadMoreChatMessages(this.graph, this._nextLink);

    this._nextLink = result.nextLink;
    this.notifyStateChange((draft: GraphChatClient) => {
      const nextMessages = result.value
        // trying to filter out system messages on the graph request causes a 400
        // delted messages are returned as messages with no content, which we can't filter on the graph request
        // so we filter them out here
        .filter(m => m.messageType === 'message' && m.body?.content)
        .map(m => graphChatMessageToACSChatMessage(m, this._userId));
      draft.messages = nextMessages.concat(draft.messages as ACSChatMessage[]).sort(MessageCreatedComparator);
      draft.onLoadPreviousChatMessages = this._nextLink ? this.loadMoreMessages : undefined;
    });
    // return true when there are no more messages to load
    return !Boolean(this._nextLink);
  };

  /**
   * Send a message to the chat thread
   *
   * @param {string} content - the content of the message
   * @memberof StatefulGraphChatClient
   */
  public sendMessage = async (content: string) => {
    if (!content) return;

    const pendingId = uuid();

    // add a pending message to the state.
    this.notifyStateChange((draft: GraphChatClient) => {
      const pendingMessage: ACSChatMessage = {
        clientMessageId: pendingId,
        messageId: pendingId,
        contentType: 'text',
        messageType: 'chat',
        content,
        senderDisplayName: this._userDisplayName,
        createdOn: new Date(),
        senderId: this._userId,
        mine: true,
        status: 'sending'
      };
      draft.messages.push(pendingMessage);
    });
    try {
      // send message
      const chat: ChatMessage = await sendChatMessage(this.graph, this._chatId, content);
      // emit new state
      this.notifyStateChange((draft: GraphChatClient) => {
        const draftIndex = draft.messages.findIndex(m => m.messageId === pendingId);
        draft.messages.splice(draftIndex, 1, graphChatMessageToACSChatMessage(chat, this._userId));
      });
    } catch (e) {
      this.notifyStateChange((draft: GraphChatClient) => {
        const draftMessage = draft.messages.find(m => m.messageId === pendingId);
        (draftMessage as ACSChatMessage).status = 'failed';
      });
      throw new Error('Failed to send message');
    }
  };

  /*
   * Helper method to set the content of a message to show deletion
   */
  private setDeletedContent = (message: ACSChatMessage) => {
    message.content = '<em>This message has been deleted.</em>';
    message.contentType = 'html';
  };

  /**
   * Handler to delete a message
   *
   * @param messageId id of the message to be deleted, this is the clientMessageId when triggered by the re-send action on a failed message, or the messageId when triggered by the delete action on a sent message
   * @returns {Promise<void>}
   */
  public deleteMessage = async (messageId: string): Promise<void> => {
    if (!messageId) return;
    const message = this._state.messages.find(m => m.messageId === messageId) as ACSChatMessage;
    // only messages not persisted to graph should have a clientMessageId
    const uncommitted = this._state.messages.find(
      m => (m as ACSChatMessage).clientMessageId === messageId
    ) as ACSChatMessage;
    if (message?.mine) {
      try {
        // uncommitted messages are not persisted to the graph, so don't call graph when deleting them
        if (!uncommitted) {
          await deleteChatMessage(this.graph, this._chatId, messageId);
        }
        this.notifyStateChange((draft: GraphChatClient) => {
          const draftMessage = draft.messages.find(m => m.messageId === messageId) as ACSChatMessage;
          if (draftMessage.clientMessageId) {
            // just remove messages that were not saved to the graph
            draft.messages.splice(draft.messages.indexOf(draftMessage), 1);
          } else {
            // show deleted messages which have been persisted to the graph as deleted in the UI
            this.setDeletedContent(draftMessage);
          }
        });
      } catch (e) {
        // TODO: How do we handle failed deletes?
      }
    }
  };

  /**
   * Update a message in the thread
   *
   * @param {string} messageId Id of the message to be updated
   * @param {string} content new content of the message
   * @memberof StatefulGraphChatClient
   */
  public updateMessage = async (messageId: string, content: string) => {
    if (!messageId || !content) return;
    const message = this._state.messages.find(m => m.messageId === messageId) as ACSChatMessage;
    if (message?.mine && message.content) {
      this.notifyStateChange((draft: GraphChatClient) => {
        const updating = draft.messages.find(m => m.messageId === messageId) as ACSChatMessage;
        if (updating) {
          updating.content = content;
          updating.status = 'sending';
        }
      });
      try {
        await updateChatMessage(this.graph, this._chatId, messageId, content);
        this.notifyStateChange((draft: GraphChatClient) => {
          const updated = draft.messages.find(m => m.messageId === messageId) as ACSChatMessage;
          updated.status = 'delivered';
        });
      } catch (e) {
        this.notifyStateChange((draft: GraphChatClient) => {
          const updating = draft.messages.find(m => m.messageId === messageId) as ACSChatMessage;
          updating.status = 'failed';
        });
        throw new Error('Failed to update message');
      }
    }
  };

  /*
   * Event handler to be called when a new message is received by the notification service
   */
  private onMessageReceived = (message: ACSChatMessage) => {
    this.notifyStateChange((draft: GraphChatClient) => {
      const index = draft.messages.findIndex(m => m.messageId === message.messageId);
      // this message is not already in thread so just add it
      if (index === -1) {
        // sort to ensure that messages are in the correct order should we get messages out of order
        draft.messages = draft.messages.concat(message).sort(MessageCreatedComparator);
      } else {
        // replace the existing version of the message with the new one
        draft.messages.splice(index, 1, message);
      }
    });
  };

  /*
   * Event handler to be called when a message deletion is received by the notification service
   */
  private onMessageDeleted = (message: ACSChatMessage) => {
    this.notifyStateChange((draft: GraphChatClient) => {
      const draftMessage = draft.messages.find(m => m.messageId === message.messageId) as ACSChatMessage;
      if (draftMessage) this.setDeletedContent(draftMessage);
    });
  };

  /*
   * Event handler to be called when a message edit is received by the notification service
   */
  private onMessageEdited = (message: ACSChatMessage) => {
    this.notifyStateChange((draft: GraphChatClient) => {
      const index = draft.messages.findIndex(m => m.messageId === message.messageId);
      draft.messages.splice(index, 1, message);
    });
  };

  private onChatNotificationsSubscribed = (messagesResource: string): void => {
    if (messagesResource.includes(`/${this._chatId}/`)) {
      void this.loadChatData();
    } else {
      // better clean this up as we don't want to be listening to events for other chats
    }
  };

  /**
   * Register event listeners for chat events to be triggered from the notification service
   */
  private registerEventListeners() {
    this._eventEmitter.on('chatMessageReceived', this.onMessageReceived);
    this._eventEmitter.on('chatMessageDeleted', this.onMessageDeleted);
    this._eventEmitter.on('chatMessageEdited', this.onMessageEdited);
    this._eventEmitter.on('chatMessageNotificationsSubscribed', this.onChatNotificationsSubscribed);
  }

  /**
   * Provided the graph instance for the component with the correct SDK version decoration
   *
   * @readonly
   * @private
   * @type {IGraph}
   * @memberof StatefulGraphChatClient
   */
  private get graph(): IGraph {
    return Providers.globalProvider.graph.forComponent('mgt-chat');
  }

  /**
   * State of the chat client with initial values set
   *
   * @private
   * @type {GraphChatClient}
   * @memberof StatefulGraphChatClient
   */
  private _state: GraphChatClient = {
    status: 'initial',
    userId: '',
    messages: [],
    participantCount: 0,
    disableEditing: false,
    numberOfChatMessagesToReload: this._messagesPerCall,
    onDeleteMessage: this.deleteMessage,
    onSendMessage: this.sendMessage,
    onUpdateMessage: this.updateMessage,
    activeErrorMessages: [],
    chat: this._chat
  };
}

export { StatefulGraphChatClient };
