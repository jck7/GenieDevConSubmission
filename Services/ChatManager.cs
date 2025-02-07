using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ExcelGenie.Models;

namespace ExcelGenie.Services
{
    public class ChatManager
    {
        private readonly List<(string message, bool isUser)> conversationHistory;
        private readonly Stack<UndoActionInfo> undoStack;
        private string? customInstructions;

        public event EventHandler<(string message, bool isUser)>? MessageAdded;
        public event EventHandler<LoadingMessageInfo>? LoadingMessageAdded;
        public event EventHandler<LoadingMessageInfo>? LoadingMessageRemoved;
        public event EventHandler<(string description, string code)>? SystemResponseAdded;
        public event EventHandler<int>? ConversationReverted;

        public ChatManager()
        {
            conversationHistory = new List<(string message, bool isUser)>();
            undoStack = new Stack<UndoActionInfo>();
            customInstructions = Properties.Settings.Default.CustomInstructions;
        }

        public void SetCustomInstructions(string instructions)
        {
            customInstructions = instructions;
            Properties.Settings.Default.CustomInstructions = instructions;
            Properties.Settings.Default.Save();
        }

        public string? GetCustomInstructions()
        {
            return customInstructions;
        }

        public void AddUserMessage(string message)
        {
            // If there are custom instructions, prepend them to the user's message in the history
            if (!string.IsNullOrEmpty(customInstructions))
            {
                string fullMessage = $"[Custom Instructions: {customInstructions}]\n{message}";
                conversationHistory.Add((fullMessage, true));
            }
            else
            {
                conversationHistory.Add((message, true));
            }
            
            // Always show the original message in the UI
            MessageAdded?.Invoke(this, (message, true));
        }

        public void AddSystemMessage(string message, LoadingMessageInfo? loadingInfo = null)
        {
            // If we have a loading message, update it instead of adding a new one
            if (loadingInfo != null && loadingInfo.HistoryIndex < conversationHistory.Count)
            {
                conversationHistory[loadingInfo.HistoryIndex] = (message, false);
                MessageAdded?.Invoke(this, (message, false));
            }
            else
            {
                conversationHistory.Add((message, false));
                MessageAdded?.Invoke(this, (message, false));
            }
        }

        public LoadingMessageInfo AddLoadingMessage()
        {
            // Add to conversation history
            conversationHistory.Add(("Generating...", false));
            int historyIndex = conversationHistory.Count - 1;

            var loadingInfo = new LoadingMessageInfo
            {
                HistoryIndex = historyIndex
            };

            LoadingMessageAdded?.Invoke(this, loadingInfo);
            return loadingInfo;
        }

        public void RemoveLoadingMessage(LoadingMessageInfo loadingInfo)
        {
            if (loadingInfo.HistoryIndex < conversationHistory.Count)
            {
                conversationHistory.RemoveAt(loadingInfo.HistoryIndex);
            }
            LoadingMessageRemoved?.Invoke(this, loadingInfo);
        }

        public void AddSystemResponse(string description, string code)
        {
            // Add the description to conversation history
            conversationHistory.Add((description, false));
            
            // Send the response
            SystemResponseAdded?.Invoke(this, (description, code));
        }

        public void PushUndoAction(UndoActionInfo action)
        {
            action.ConversationHistoryCount = conversationHistory.Count;
            undoStack.Push(action);
        }

        public UndoActionInfo? PopUndoAction()
        {
            if (undoStack.Count == 0)
            {
                AddSystemMessage("No actions to undo.");
                return null;
            }

            var lastAction = undoStack.Pop();

            // Remove messages from conversationHistory after revertCount
            while (conversationHistory.Count > lastAction.ConversationHistoryCount)
            {
                conversationHistory.RemoveAt(conversationHistory.Count - 1);
            }

            ConversationReverted?.Invoke(this, lastAction.ChatPanelChildrenCount);
            return lastAction;
        }

        public List<(string message, bool isUser)> GetConversationHistory()
        {
            return new List<(string message, bool isUser)>(conversationHistory);
        }

        public void ClearConversation()
        {
            conversationHistory.Clear();
            undoStack.Clear();
        }

        public bool HasUndoActions => undoStack.Count > 0;
    }
} 