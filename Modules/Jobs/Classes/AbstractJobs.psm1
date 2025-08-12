class AbstractJobs {

    $sleep
    $max

    AbstractJobs([int]$max, [int]$sleep) {   
        $this.sleep = $sleep
        $this.max = $max
    }   

 
    [hashtable]ProcessCompleted() {
        throw "The method 'ProcessCompleted' must be implemented in a derived class."
    }

    [void]Throttle() {
        # we only want a max of 7 running jobs.  Wait until some complete
        $runningJobs = Get-Job -State Running
        while ($runningJobs.Count -gt $this.max) {
            Sleep ($this.sleep);
            $COUNT = $runningJobs.Count;
            Write ("Running jobs is maxed at $COUNT");
            $runningJobs = get-job -State Running;
        }
    }

    [void]WaitForJobsToComplete() {
        #finish up remaining jobs
        $runningJobs = get-job -State Running
        while ($runningJobs.Count -gt 0) {
            Sleep (5);
            $runningJobs = get-job -State Running;
        }
    }
}


