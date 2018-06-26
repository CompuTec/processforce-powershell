Class CTTimer {
    [DateTime] $startDateTime;
    [DateTime] $endDateTime;
    [DateTime] $prevStepDateTime;
    [DateTime] $currentDateTime;
    
    CTTimer() {
        $this.startDateTime = Get-Date;
        $this.prevStepDateTime = $this.startDateTime;
    }

    [double] round(){
        $this.currentDateTime = Get-Date;
        $seconds = ($this.currentDateTime-$this.prevStepDateTime).TotalSeconds;
        $this.prevStepDateTime = $this.currentDateTime;
        return $seconds;
    }

    [double] totalSeconds(){
        $this.endDateTime =  Get-Date;
        $seconds = ($this.startDateTime-$this.endDateTime).TotalSeconds;
        return $seconds;
    }

    [void] stop(){
        $this.endDateTime =  Get-Date;
    }

    [void] restart() {
        $this.startDateTime = Get-Date;
        $this.prevStepDateTime = $this.startDateTime;
    }



}
    