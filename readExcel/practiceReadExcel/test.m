t=timer

t.StartDelay = 1;%延时1秒开始

t.ExecutionMode = 'fixedRate';%启用循环执行

t.Period = 1;%循环间隔1秒

t.TasksToExecute =50;%循环次数5次

t.TimerFcn = @WriteInExcel;%开始执行

start(t)