using System;
using System.Collections.Concurrent;
using Autodesk.Revit.UI;

namespace THBIM.Services
{
    /// <summary>
    /// Generic external event handler that executes queued actions
    /// inside a valid Revit API context. Required for write operations
    /// (Transaction) from modeless windows.
    /// </summary>
    internal sealed class RevitEventHandler : IExternalEventHandler
    {
        private static RevitEventHandler _instance;
        private static ExternalEvent _externalEvent;

        private readonly ConcurrentQueue<Action<UIApplication>> _actions = new();
        private Action _onComplete;

        public static RevitEventHandler Instance => _instance;
        public static ExternalEvent Event => _externalEvent;

        /// <summary>
        /// Initialize once (call from IExternalCommand.Execute or App.OnStartup).
        /// </summary>
        public static void Initialize()
        {
            if (_instance != null) return;
            _instance = new RevitEventHandler();
            _externalEvent = ExternalEvent.Create(_instance);
        }

        /// <summary>
        /// Queue an action to run inside the Revit API context, then raise the event.
        /// </summary>
        public void Enqueue(Action<UIApplication> action, Action onComplete = null)
        {
            if (action == null) return;
            _onComplete = onComplete;
            _actions.Enqueue(action);
            _externalEvent?.Raise();
        }

        public void Execute(UIApplication app)
        {
            while (_actions.TryDequeue(out var action))
            {
                try
                {
                    action(app);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"RevitEventHandler error: {ex.Message}");
                }
            }

            try
            {
                _onComplete?.Invoke();
            }
            catch { }
        }

        public string GetName() => "SheetLink.RevitEventHandler";
    }
}
