using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using ExcelDna.Integration;

namespace TestAsyncError
{
    public static class Functions
    {
        static HashSet<AsyncCallInfo> _errorCalls = new HashSet<AsyncCallInfo>();

        public static object RunClock(string input, bool fail)
        {
            // We build an AsyncCallInfo that is used to identify calls to this function
            // We need to have the same information in callInfo as is passed to the ExcelAsyncUtil.Observe call
            var functionName = nameof(RunClock);
            var args = new object[] { input, fail };
            var callInfo = new AsyncCallInfo(functionName, args); // This will be the key in our _errorCalls set

            // We track whether a call to this function (with these parameters) is in the error state in the _errorCalls HashSet
            if (!_errorCalls.Contains(callInfo))
            {
                // This is where we start in the normal case - we're not in the error state - call the observable as usual
                var result = ExcelAsyncUtil.Observe(functionName, args, () => new ExcelObservableClock(input, fail));

                // Typically we expect the first result here to be #N/A, though that is not always so
                // We need to detect whether the result indicates an error, and if so change to error handling mode
                if (!result.Equals(ExcelError.ExcelErrorValue))
                {
                    // We return the result of the observable call as is
                    return result;
                }
                else
                {
                    // We've detected an error from the observable, so we switch to error handling mode
                    _errorCalls.Add(callInfo);

                    // Then call the error fallback function
                    // (arguments sent to the error callback need not be the same as the original call)
                    return XlCall.Excel(XlCall.xlUDF, nameof(AsyncErrorFallback), input);
                }
            }
            else
            {
                // If we are in the error handling state, we need to continue calling the error fallback
                // until it returns a value other than #N/A
                var errorFallbackResult = AsyncErrorFallback(input);
                if (!errorFallbackResult.Equals(ExcelError.ExcelErrorNA))
                {
                    // We have a result from the error fallback - remove the cached error
                    // which means the next call to the function will restart the original observable
                    _errorCalls.Remove(callInfo);
                }
                // Return the result of the error fallback
                return errorFallbackResult;
            }
        }

        // This is the error fallback function - it is called when the observable returns an error
        // It can be any other UDF that internally calls RTD
        // It will continue to be the fallback function until it returns a non-#N/A value
        public static object AsyncErrorFallback(object input)
        {
            var functionName = nameof(AsyncErrorFallback);
            var args = new object[] { input };
            var result = ExcelAsyncUtil.Run(functionName, args, () =>
            {
                Thread.Sleep(1000);
                return $"Error result at {DateTime.Now:T} for {input}";
            });

            return result;
        }
    }

    // This is just a sample observable that returns the current time every second.
    // It takes a flag to indicate whether it should fail (throw an exception) or not.
    class ExcelObservableClock : IExcelObservable
    {
        Timer _timer;
        IExcelObserver _observer;
        string _param;
        bool _fail;

        public ExcelObservableClock(string param, bool fail)
        {
            Debug.WriteLine("Created " + param);
            _param = param;
            _timer = new Timer(timer_tick, null, 0, 1000);
            _fail = fail;
        }

        public IDisposable Subscribe(IExcelObserver observer)
        {
            _observer = observer;
            // observer.OnNext(DateTime.Now.ToString("HH:mm:ss.fff") + " (Subscribe)");
            return new ActionDisposable(() =>
            {
                _observer = null;
                Debug.WriteLine("Disposed " + _param);
            });
        }

        void timer_tick(object _)
        {
            if (_fail)
            {
                _observer?.OnError(new Exception("Error at " + DateTime.Now.ToString("HH:mm:ss.fff")));
                return;
            }
            string now = DateTime.Now.ToString("HH:mm:ss.fff");
            _observer?.OnNext(now);
        }

        class ActionDisposable : IDisposable
        {
            readonly Action _disposeAction;
            public ActionDisposable(Action disposeAction)
            {
                _disposeAction = disposeAction;
            }
            public void Dispose()
            {
                _disposeAction();
            }
        }
    }

    // Encapsulates the information that defines and async call or observable hook-up.
    // Checked for equality and stored in a Dictionary, so we have to be careful
    // to define value equality and a consistent HashCode.

    // Used as Keys in a Dictionary - should be immutable. 
    // We allow parameters to be null or primitives or ExcelReference objects, 
    // or a 1D array or 2D array with primitives or arrays.
    internal struct AsyncCallInfo : IEquatable<AsyncCallInfo>
    {
        readonly string _functionName;
        readonly object _parameters;
        readonly int _hashCode;

        public AsyncCallInfo(string functionName, object parameters)
        {
            _functionName = functionName;
            _parameters = parameters;
            _hashCode = 0; // Need to set to some value before we call a method.
            _hashCode = ComputeHashCode();
        }

        // Jon Skeet: http://stackoverflow.com/questions/263400/what-is-the-best-algorithm-for-an-overridden-system-object-gethashcode
        int ComputeHashCode()
        {
            unchecked
            {
                int hash = 17;
                hash = hash * 23 + (_functionName == null ? 0 : _functionName.GetHashCode());
                hash = hash * 23 + ComputeHashCode(_parameters);
                return hash;
            }
        }

        // Computes a hash code for the parameters, consistent with the value equality that we define.
        // Also checks that the data types passed for parameters are among those we handle properly for value equality.
        // For now no string[]. But invalid types passed in will causes exception immediately.
        static int ComputeHashCode(object obj)
        {
            if (obj == null) return 0;

            // CONSIDER: All of this could be replaced by a check for (obj is ValueType || obj is ExcelReference)
            //           which would allow a more flexible set of parameters, at the risk of comparisons going wrong.
            //           We can reconsider if this arises, or when we implement async automatically or custom marshaling 
            //           to other data types. For now this allows everything that can be passed as parameters from Excel-DNA.

            // We also support using an opaque byte[] hash as the parameters 'key'.
            // In cases with huge amounts of active topics, especially using string parameters, this can improve the memory usage significantly.

            if (obj is double ||
                obj is float ||
                obj is string ||
                obj is bool ||
                obj is DateTime ||
                obj is ExcelReference ||
                obj is ExcelError ||
                obj is ExcelEmpty ||
                obj is ExcelMissing ||
                obj is int ||
                obj is uint ||
                obj is long ||
                obj is ulong ||
                obj is short ||
                obj is ushort ||
                obj is byte ||
                obj is sbyte ||
                obj is decimal ||
                obj.GetType().IsEnum)
            {
                return obj.GetHashCode();
            }

            unchecked
            {
                int hash = 17;

                double[] doubles = obj as double[];
                if (doubles != null)
                {
                    foreach (double item in doubles)
                    {
                        hash = hash * 23 + item.GetHashCode();
                    }
                    return hash;
                }

                double[,] doubles2 = obj as double[,];
                if (doubles2 != null)
                {
                    foreach (double item in doubles2)
                    {
                        hash = hash * 23 + item.GetHashCode();
                    }
                    return hash;
                }

                object[] objects = obj as object[];
                if (objects != null)
                {
                    foreach (object item in objects)
                    {
                        hash = hash * 23 + ((item == null) ? 0 : ComputeHashCode(item));
                    }
                    return hash;
                }

                object[,] objects2 = obj as object[,];
                if (objects2 != null)
                {
                    foreach (object item in objects2)
                    {
                        hash = hash * 23 + ((item == null) ? 0 : ComputeHashCode(item));
                    }
                    return hash;
                }

                byte[] bytes = obj as byte[];
                if (bytes != null)
                {
                    foreach (byte b in bytes)
                    {
                        hash = hash * 23 + b;
                    }
                    return hash;
                }
            }
            throw new ArgumentException("Invalid type used for async parameter(s)", "parameters");
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (obj.GetType() != typeof(AsyncCallInfo)) return false;
            return Equals((AsyncCallInfo)obj);
        }

        public bool Equals(AsyncCallInfo other)
        {
            if (_hashCode != other._hashCode) return false;
            return Equals(other._functionName, _functionName)
                   && ValueEquals(_parameters, other._parameters);
        }

        #region Helpers to implement value equality
        // The value equality we check here is for the types we allow in CheckParameterTypes above.
        static bool ValueEquals(object a, object b)
        {
            if (Equals(a, b)) return true; // Includes check for both null
            if (a is double[] && b is double[]) return ArrayEquals((double[])a, (double[])b);
            if (a is double[,] && b is double[,]) return ArrayEquals((double[,])a, (double[,])b);
            if (a is object[] && b is object[]) return ArrayEquals((object[])a, (object[])b);
            if (a is object[,] && b is object[,]) return ArrayEquals((object[,])a, (object[,])b);
            if (a is byte[] && b is byte[]) return ArrayEquals((byte[])a, (byte[])b);
            return false;
        }

        static bool ArrayEquals(double[] a, double[] b)
        {
            if (a.Length != b.Length)
                return false;
            for (int i = 0; i < a.Length; i++)
            {
                if (a[i] != b[i]) return false;
            }
            return true;
        }

        static bool ArrayEquals(double[,] a, double[,] b)
        {
            int rows = a.GetLength(0);
            int cols = a.GetLength(1);
            if (rows != b.GetLength(0) ||
                cols != b.GetLength(1))
            {
                return false;
            }
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    if (a[i, j] != b[i, j]) return false;
                }
            }
            return true;
        }

        static bool ArrayEquals(object[] a, object[] b)
        {
            if (a.Length != b.Length)
                return false;
            for (int i = 0; i < a.Length; i++)
            {
                if (!ValueEquals(a[i], b[i]))
                    return false;
            }
            return true;
        }

        static bool ArrayEquals(object[,] a, object[,] b)
        {
            int rows = a.GetLength(0);
            int cols = a.GetLength(1);
            if (rows != b.GetLength(0) ||
                cols != b.GetLength(1))
            {
                return false;
            }
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    if (!ValueEquals(a[i, j], b[i, j]))
                        return false;
                }
            }
            return true;
        }

        static bool ArrayEquals(byte[] a, byte[] b)
        {
            if (a.Length != b.Length)
                return false;
            for (int i = 0; i < a.Length; i++)
            {
                if (a[i] != b[i]) return false;
            }
            return true;
        }

        #endregion

        public override int GetHashCode()
        {
            return _hashCode;
        }

        public static bool operator ==(AsyncCallInfo asyncCallInfo1, AsyncCallInfo asyncCallInfo2)
        {
            return asyncCallInfo1.Equals(asyncCallInfo2);
        }

        public static bool operator !=(AsyncCallInfo asyncCallInfo1, AsyncCallInfo asyncCallInfo2)
        {
            return !(asyncCallInfo1.Equals(asyncCallInfo2));
        }
    }
}
