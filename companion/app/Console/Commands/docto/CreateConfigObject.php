<?php

namespace App\Console\Commands\docto;

use Illuminate\Console\Command;
use Illuminate\Support\Facades\Blade;
use Illuminate\Support\Facades\Storage;

class CreateConfigObject extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'make:configObject {name} {--paramlist="" : List of parameters accepted (comma separated) eg \'-F,--FileType\' }';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Create a Config obeject in DocTo Directory.';

    /**
     * Execute the console command.
     */
    public function handle()
    {
        $name = $this->argument('name');
        $paramlist = str( $this->option('paramlist'))->explode(',');
       $pasfile =  Blade::render('docto.pasfile.configObject',['paramlist'=>$paramlist, 'name'=>$name]);
       $storageDisk = Storage::build([
           'driver' => 'local',
           'root' => base_path('../src/ParamObjects'),
       ]);
       $storageDisk->put('config' . $name . '.pas', $pasfile);
       $this->info('Create New Parameter Config file. ' . 'config' . $name . '.pas' );
    }
}
