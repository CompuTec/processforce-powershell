using module .\CTTimer.psm1;
Class CTLogger {
    [CTTimer] $cttimer;
    [System.Collections.Generic.Dictionary[string,psobject]] $tasks;
    [string] $task;

    CTLogger($connectionType,$task) {
        $this.ctimer = New-Object CTTimer;
        $this.tasks = New-Object 'System.Collections.Generic.Dictionary[string,psobject]';
    }

    [void] startSubtask($taskName){

    }

    [void] endSubtTask($taskName,$status,$message) {

    }

}
    #lines
            #dbtype
            #connection type (DI API, UI API)
            #task (itemDetails import, )
            #subtask (connecting, import BOM, import ITD, import production process, open BOM form)
            #start datetime
            #end datetime
            #duration