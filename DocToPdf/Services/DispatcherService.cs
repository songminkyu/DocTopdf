
using System;
using System.Windows;
using System.Windows.Threading;
using System.Threading;
using System.Threading.Tasks;
using DocToPdf.Services;

namespace DocToPdf.UIControlServices
{       
    public class DispatcherService
    {
        /*
         NotSupportedException
         이 형식의 CollectionView에서는 발송자 스레드와 다른 스레드에서의 해당 SourceCollection에 대한 변경 내용을 지원하지 않습니다.
         -----
         UI 쓰레드에 의해 점유중인 자원에 대해선 자원에 수정을 가할 때 UI 쓰레드의 Dispatcher에 작업을 위임해야 한다. 

         <호출 방법>
         DispatcherHelper.Invoke((System.Action)(( ) =>
         {
             // your logic
         }));
         */        
        public static bool OnUpdateWindow { get; set; } = false;
        #region Invoke
        public static void Invoke(Action action)
        {          
            try
            {
                if (Application.Current == null)
                    return;

                bool IsRequestedFromMainThread = Application.Current.Dispatcher.CheckAccess();

                if (IsRequestedFromMainThread == true)
                {                    
                    action.Invoke();                  
                } 
                else
                {
                    Application.Current.Dispatcher.Invoke(DispatcherPriority.Normal, action);                    
                }
                
            }
            catch (Exception ex)
            {
                LoggingService.Logger("DispatcherService/Normal : " + ex.Message, LogLevel.Error);
            }
        }
        
        public static void InvokeBackground(Action action)
        {           
            try
            {
                if (Application.Current == null)
                    return;

                bool IsRequestedFromMainThread = Application.Current.Dispatcher.CheckAccess();

                if (IsRequestedFromMainThread == true)
                {                  
                    action.Invoke();                    
                }
                else
                {
                    Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, action);
                }
                
            }
            catch (Exception ex)
            {
                LoggingService.Logger("DispatcherService/Background : " + ex.Message, LogLevel.Error);                
            }
        }
        public static TResult InvokeBackground<TResult>(Func<TResult> func) where TResult : class
        {
            TResult result = null;

            try
            {
                if (Application.Current == null)
                    return result;

                bool IsRequestedFromMainThread = Application.Current.Dispatcher.CheckAccess();

                if (IsRequestedFromMainThread == true)
                {
                    result = func.Invoke();
                }
                else
                {
                    result = Application.Current.Dispatcher.Invoke(func, DispatcherPriority.Background);
                }
            }
            catch (Exception ex)
            {
                LoggingService.Logger("DispatcherService/InvokeBackground : " + ex.Message, LogLevel.Error);                
            }
            return result;
        }
        public static void InvokeIdle(Action action)
        {            
            try
            {
                if (Application.Current == null)
                    return;

                bool IsRequestedFromMainThread = Application.Current.Dispatcher.CheckAccess();

                if (IsRequestedFromMainThread == true)
                {
                    action.Invoke();
                }
                else
                {
                    Application.Current.Dispatcher.Invoke(DispatcherPriority.ApplicationIdle, action);
                }                
            }
            catch (Exception ex)
            {
                LoggingService.Logger("DispatcherService/InvokeBackground : " + ex.Message, LogLevel.Error);                
            }
        }
        #endregion
        #region BeginInvoke
        public static async Task<bool> BeginInvokeNormal(Action action)
        {                              
            try
            {
                if (Application.Current == null)
                    return false;

                bool IsRequestedFromMainThread = Application.Current.Dispatcher.CheckAccess();

                if (IsRequestedFromMainThread == true)
                {                   
                    action.Invoke();
                }
                else
                {                                                          
                    await Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Normal, action);
                }

            }
            catch (Exception ex)
            {                
                LoggingService.Logger("DispatcherService/BeginNormal : " + ex.Message, LogLevel.Error);
            }          
            return true;
        }
        public static async Task<bool> BeginInvokeBackground(Action action)
        {         
            try
            {                
                if (Application.Current == null)
                    return false;
            
                bool IsRequestedFromMainThread = Application.Current.Dispatcher.CheckAccess();

                if (IsRequestedFromMainThread == true)
                {                    
                    action.Invoke();                   
                }
                else
                {
                    await Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Background, action);                    
                }
            }
            catch (Exception ex)
            {                              
                LoggingService.Logger("DispatcherService/BeginBackground : " + ex.Message, LogLevel.Error);
            }                     
            return true;
        }
        #endregion
        #region InvokeAsync
        public static async Task<bool> InvokeAsyncNormal(Action action)
        {
            try
            {
                if (Application.Current == null)
                    return false;

                bool IsRequestedFromMainThread = Application.Current.Dispatcher.CheckAccess();

                if (IsRequestedFromMainThread == true)
                {
                    action.Invoke();
                }
                else
                {
                    await Application.Current.Dispatcher.InvokeAsync(action, DispatcherPriority.Normal);
                }

            }
            catch (Exception ex)
            {
                LoggingService.Logger("DispatcherService/BeginNormal : " + ex.Message, LogLevel.Error);                
            }
            return true;
        }
        public static async Task<bool> InvokeAsyncBackground(Action action)
        {
            try
            {
                if (Application.Current == null)
                    return false;

                bool IsRequestedFromMainThread = Application.Current.Dispatcher.CheckAccess();

                if (IsRequestedFromMainThread == true)
                {
                    action.Invoke();
                }
                else
                {
                    await Application.Current.Dispatcher.InvokeAsync(action, DispatcherPriority.Background);
                }

            }
            catch (Exception ex)
            {
                LoggingService.Logger("DispatcherService/BeginNormal : " + ex.Message, LogLevel.Error);                
            }
            return true;
        }
        public static async Task<TResult> InvokeAsyncNormal<TResult>(Func<TResult> func) where TResult : class, new()
        {
            TResult result = new TResult();

            try
            {
                if (Application.Current == null)
                    return result;

                bool IsRequestedFromMainThread = Application.Current.Dispatcher.CheckAccess();

                if (IsRequestedFromMainThread == true)
                {
                    result = func.Invoke();
                }
                else
                {
                    result = await Application.Current.Dispatcher.InvokeAsync(func, DispatcherPriority.Normal);
                }
            }
            catch (Exception ex)
            {
                LoggingService.Logger("DispatcherService/BeginNormal : " + ex.Message, LogLevel.Error);                
            }
            return result;
        }
        public static async Task<TResult> InvokeAsyncBackground<TResult>(Func<TResult> func) where TResult : class, new()
        {
            TResult result = new TResult();

            try
            {
                if (Application.Current == null)
                    return result;

                bool IsRequestedFromMainThread = Application.Current.Dispatcher.CheckAccess();

                if (IsRequestedFromMainThread == true)
                {
                    result = func.Invoke();
                }
                else
                {
                    result = await Application.Current.Dispatcher.InvokeAsync(func, DispatcherPriority.Background);
                }
            }
            catch (Exception ex)
            {
                LoggingService.Logger("DispatcherService/InvokeAsyncBackground : " + ex.Message, LogLevel.Error);                
            }
            return result;
        }
        
        #endregion

        #region 이벤트 실행하기 - DoEvents()
        /// <summary>
        /// 이벤트 실행하기
        /// </summary>        
        private static readonly DispatcherOperationCallback ExitFrameCallback = ExitFrame;
        public static void DoEvents(DispatcherPriority priority = DispatcherPriority.Background)
        {
            var nestedFrame = new DispatcherFrame();

            var exitOperation = Dispatcher.CurrentDispatcher.BeginInvoke(priority, ExitFrameCallback, nestedFrame);

            try
            {                
                Dispatcher.PushFrame(nestedFrame);
                
                if (exitOperation.Status != DispatcherOperationStatus.Completed)
                    exitOperation.Abort();
            }
            catch
            {
                exitOperation.Abort();
            }
        }
        #endregion

        #region 프레임 종료하기 - ExitFrame(frame)
        /// <summary>
        /// 프레임 종료하기
        /// </summary>
        /// <param name="frame">프레임</param>
        /// <returns>처리 결과</returns>
        private static object? ExitFrame(object frame)
        {
            ((DispatcherFrame)frame).Continue = false;

            return null;
        }
        #endregion
       
        #region Set Thread Culture
        public static void SetThreadLanguage(Thread CurrentThread, string CurrentLangCode)
        {
            CurrentThread.CurrentCulture = new System.Globalization.CultureInfo(CurrentLangCode);
            CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo(CurrentLangCode);
        }
        #endregion
    
    }
}
