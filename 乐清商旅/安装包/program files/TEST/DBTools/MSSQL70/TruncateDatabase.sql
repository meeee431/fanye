
/***************************************/
/*                                     */
/*         Êý¾Ý¿âËõÐ¡                    */
/*                                     */
/***************************************/

BEGIN TRANSACTION            
   DECLARE @ReturnCode INT    
   SELECT @ReturnCode = 0     
   DECLARE @JobID BINARY(16)  

   DECLARE @CmdStr VARCHAR(8000)

   IF (SELECT COUNT(*) FROM msdb.dbo.syscategories WHERE name = N'[Uncategorized (Local)]') < 1 
     EXECUTE msdb.dbo.sp_add_category @name = N'[Uncategorized (Local)]'

   SELECT @JobID = job_id FROM   msdb.dbo.sysjobs WHERE (name = N'TruncateDatabase')       
   IF (@JobID IS NOT NULL)    
     GOTO  EndSave              

   SELECT @CmdStr='EXEC TruncateDatabase'

  EXECUTE @ReturnCode = msdb.dbo.sp_add_job 
      @job_id = @JobID OUTPUT , 
      @job_name = N'TruncateDatabase', 
      @owner_login_name = N'sa', 
      @description = N'No description available.', 
      @category_name = N'[Uncategorized (Local)]', 
      @enabled = 1, 
      @notify_level_email = 0, 
      @notify_level_page = 0, 
      @notify_level_netsend = 0, 
      @notify_level_eventlog = 2, 
      @delete_level= 0
  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 

  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep 
      @job_id = @JobID, 
      @step_id = 1, 
      @step_name = N'TruncateDatabase', 
      @command = @CmdStr, 
      @database_name = N'[Database]', 
      @server = N'[Server]', 
      @database_user_name = N'sa', 
      @subsystem = N'TSQL', 
      @cmdexec_success_code = 0, 
      @flags = 4, 
      @retry_attempts = 3, 
      @retry_interval = 5, 
      @output_file_name = N'', 
      @on_success_step_id = 0, 
      @on_success_action = 1, 
      @on_fail_step_id = 0, 
      @on_fail_action = 2
  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 

  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id = @JobID, @name = N'TruncateDatabase', @enabled = 1, @freq_type = 4, @active_start_date = 20000101, @active_start_time = 010000, @freq_interval = 1, @freq_subday_type = 1, @freq_subday_interval = 0, @freq_relative_interval = 0, @freq_recurrence_factor = 0, @active_end_date = 99991231, @active_end_time = 235959
  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 

  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @JobID, @server_name = N'(local)' 
  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 

  COMMIT TRANSACTION          
  GOTO   EndSave              
  QuitWithRollback:
  IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION 
  EndSave: 

GO
