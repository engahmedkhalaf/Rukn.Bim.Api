using System;
using System.Collections.Concurrent;
using Autodesk.Revit.UI;

namespace QicBoqMapper
{
    public class QicBoqExternalEventHandler : IExternalEventHandler
    {
        private readonly ConcurrentQueue<Action<UIApplication>> _actions = new ConcurrentQueue<Action<UIApplication>>();

        public void QueueAction(Action<UIApplication> action)
        {
            _actions.Enqueue(action);
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
                    TaskDialog.Show("External Event Error", $"Failed to execute operation:\n{ex.Message}");
                }
            }
        }

        public string GetName()
        {
            return "QIC BOQ Modeless Event Handler";
        }
    }
}
