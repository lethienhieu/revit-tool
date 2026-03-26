using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace THBIM.Helpers
{
    /// <summary>
    /// Wrapper an toàn cho Revit Transaction.
    /// Tự động RollBack nếu có Exception.
    /// Hỗ trợ gộp nhiều bước vào 1 transaction (batch).
    /// </summary>
    public static class TransactionHelper
    {
        // ══════════════════════════════════════════════════════════════════
        // SINGLE TRANSACTION
        // ══════════════════════════════════════════════════════════════════

        /// <summary>
        /// Thực thi action trong 1 Transaction. Tự RollBack nếu lỗi.
        /// </summary>
        /// <returns>True = commit thành công.</returns>
        public static bool Run(Document doc, string name,
                                Action<Transaction> action,
                                Action<Exception> onError = null)
        {
            using var tx = new Transaction(doc, name);
            try
            {
                tx.Start();
                action(tx);
                tx.Commit();
                return true;
            }
            catch (Exception ex)
            {
                if (tx.GetStatus() == TransactionStatus.Started)
                    tx.RollBack();

                onError?.Invoke(ex);
                return false;
            }
        }

        /// <summary>
        /// Phiên bản trả về kết quả T.
        /// </summary>
        public static (bool Success, T Result, string Error)
            Run<T>(Document doc, string name, Func<Transaction, T> func)
        {
            using var tx = new Transaction(doc, name);
            try
            {
                tx.Start();
                var result = func(tx);
                tx.Commit();
                return (true, result, null);
            }
            catch (Exception ex)
            {
                if (tx.GetStatus() == TransactionStatus.Started)
                    tx.RollBack();

                return (false, default, ex.Message);
            }
        }

        // ══════════════════════════════════════════════════════════════════
        // BATCH TRANSACTION (nhiều element, rollback toàn bộ nếu lỗi)
        // ══════════════════════════════════════════════════════════════════

        /// <summary>
        /// Xử lý batch items trong 1 Transaction lớn.
        /// Nếu 1 item lỗi → bỏ qua item đó, tiếp tục (không rollback toàn bộ).
        /// Nếu lỗi nghiêm trọng → rollback tất cả.
        /// </summary>
        public static BatchResult RunBatch<T>(
            Document doc,
            string name,
            IEnumerable<T> items,
            Action<T, List<string>> processItem,
            Action<int, int> onProgress = null)
        {
            var result = new BatchResult();
            var itemList = new List<T>(items);
            int total = itemList.Count;

            using var tx = new Transaction(doc, name);
            try
            {
                tx.Start();

                for (int i = 0; i < total; i++)
                {
                    onProgress?.Invoke(i + 1, total);
                    try
                    {
                        processItem(itemList[i], result.Errors);
                        result.Processed++;
                    }
                    catch (Exception ex)
                    {
                        result.Errors.Add($"Item {i + 1}: {ex.Message}");
                    }
                }

                tx.Commit();
                result.Success = true;
            }
            catch (Exception ex)
            {
                if (tx.GetStatus() == TransactionStatus.Started)
                    tx.RollBack();

                result.Success = false;
                result.Errors.Add($"Transaction lỗi nghiêm trọng — toàn bộ bị hủy: {ex.Message}");
            }

            return result;
        }

        // ══════════════════════════════════════════════════════════════════
        // SUB-TRANSACTION
        // ══════════════════════════════════════════════════════════════════

        /// <summary>
        /// Dùng SubTransaction khi đã có Transaction cha đang mở.
        /// </summary>
        public static bool RunSub(Document doc, Action action,
                                   Action<Exception> onError = null)
        {
            using var sub = new SubTransaction(doc);
            try
            {
                sub.Start();
                action();
                sub.Commit();
                return true;
            }
            catch (Exception ex)
            {
                sub.RollBack();
                onError?.Invoke(ex);
                return false;
            }
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    // BATCH RESULT
    // ══════════════════════════════════════════════════════════════════════

    public class BatchResult
    {
        public bool Success { get; set; }
        public int Processed { get; set; }
        public List<string> Errors { get; } = new();

        public string Summary =>
            $"Processed: {Processed}" +
            (Errors.Count > 0 ? $" | Errors: {Errors.Count}" : " | No errors");
    }
}