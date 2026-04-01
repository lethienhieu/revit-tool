using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using THBIM.Helpers;
using THBIM.Models;
// using static THBIM.Views.Viewbasecompat; // Not needed for base

namespace THBIM.Views
{
    public abstract class ViewBase : UserControl
    {
        protected static readonly SolidColorBrush BgSel = new(Color.FromRgb(0xDD, 0xEE, 0xFF));
        protected static readonly SolidColorBrush BgNorm = Brushes.Transparent;
        protected static readonly SolidColorBrush BgHover = new(Color.FromRgb(0xF0, 0xF6, 0xFF));

        // ── Filter Helpers ────────────────────────────────────────────────
        protected void EnsureAtLeastOneChecked(params CheckBox[] cbs)
        {
            if (cbs.Any(cb => cb?.IsChecked == true)) return;
            foreach (var cb in cbs) cb.IsChecked = true;
        }

        protected string UpdateFilterText(CheckBox inst, CheckBox typ, CheckBox ro, ComboBox cbFilter)
        {
            var parts = new List<string>();
            if (inst?.IsChecked == true) parts.Add("Instance");
            if (typ?.IsChecked == true) parts.Add("Type");
            if (ro?.IsChecked == true) parts.Add("Read-only");
            var text = string.Join(", ", parts);
            cbFilter.Text = text;
            return text;
        }

        protected bool PassKind(ParamKind kind, CheckBox inst, CheckBox typ, CheckBox ro) 
            => kind switch
            {
                ParamKind.Instance => inst?.IsChecked == true,
                ParamKind.Type => typ?.IsChecked == true,
                ParamKind.ReadOnly => ro?.IsChecked == true,
                _ => false
            };

        // ── UI Safe LINQ ─────────────────────────────────────────────────
        protected IEnumerable<T> SafeChildren<T>(Panel panel) where T : UIElement
            => panel?.Children?.OfType<T>() ?? Enumerable.Empty<T>();

        protected Border SafeFirstBorder(Panel panel, object tag) 
            => SafeChildren<Border>(panel).FirstOrDefault(b => ReferenceEquals(b?.Tag, tag));

        // ── Param Row Helpers ────────────────────────────────────────────
        protected virtual Border CreateAvailableRow(ParameterItem p, Action<ParameterItem> moveAction)
        {
            return ParamRowHelper.CreateRow(p, moveAction);
        }

        protected virtual Border CreateSelectedRow(ParameterItem p, Action<ParameterItem> moveAction)
        {
            return ParamRowHelper.CreateRow(p, moveAction);
        }

        // ── Apply Filters (override cho specific)
        protected virtual void ApplyAvailableFilters() { /* Impl in child */ }
        protected virtual void ApplySelectedFilters() { /* Impl in child */ }

        // ── Move Params (null-safe)
        protected void MoveToSelected(Panel avail, Panel sel, ParameterItem p)
        {
            if (p == null) return;
            var existing = SafeChildren<Border>(sel).Any(b => SameParam(b?.Tag as ParameterItem, p));
            if (existing) return;
            var row = SafeChildren<Border>(avail).FirstOrDefault(b => SameParam(b?.Tag as ParameterItem, p));
            row?.Let(r => avail.Children.Remove(r));
            p.IsHighlighted = false;
            sel.Children?.Add(CreateSelectedRow(p, (pp) => MoveToAvailable(avail, sel, pp)));
            ApplyAvailableFilters();
            ApplySelectedFilters();
        }

        protected void MoveToAvailable(Panel avail, Panel sel, ParameterItem p)
        {
            if (p == null) return;
            var existing = SafeChildren<Border>(avail).Any(b => SameParam(b?.Tag as ParameterItem, p));
            if (existing) return;
            var row = SafeChildren<Border>(sel).FirstOrDefault(b => SameParam(b?.Tag as ParameterItem, p));
            row?.Let(r => sel.Children.Remove(r));
            p.IsHighlighted = false;
            avail.Children?.Add(CreateAvailableRow(p, GetMoveToSelected(avail, sel)));
            ApplyAvailableFilters();
            ApplySelectedFilters();
        }

        private Action<ParameterItem> GetMoveToSelected(Panel avail, Panel sel)
            => p => MoveToSelected(avail, sel, p);

        private Action<ParameterItem> GetMoveToAvailable(Panel avail, Panel sel)
            => p => MoveToAvailable(avail, sel, p);

        protected static bool SameParam(ParameterItem a, ParameterItem b)
            => a != null && b != null && string.Equals(a.Name, b.Name, StringComparison.OrdinalIgnoreCase);

        protected Border GetHlSel(Panel sel) 
            => SafeChildren<Border>(sel).FirstOrDefault(b => b?.Tag is ParameterItem p && p.IsHighlighted);

        // ── Sort Helpers
        protected void SortSelected(Panel sel, string orderBy = "Name")
        {
            var items = SafeChildren<Border>(sel).OrderBy(b => (b?.Tag as ParameterItem)?.Name).ToList();
            sel.Children.Clear();
            foreach (var i in items) sel.Children.Add(i);
            ApplySelectedFilters();
        }
    }

    // ── Extensions cho null-safety
    public static class ViewBaseExtensions
    {
        public static void Let<T>(this T obj, Action<T> action) where T : class
        {
            if (obj != null) action(obj);
        }
    }
}
