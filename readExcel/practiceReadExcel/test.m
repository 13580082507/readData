t=timer

t.StartDelay = 1;%��ʱ1�뿪ʼ

t.ExecutionMode = 'fixedRate';%����ѭ��ִ��

t.Period = 1;%ѭ�����1��

t.TasksToExecute =50;%ѭ������5��

t.TimerFcn = @WriteInExcel;%��ʼִ��

start(t)