<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;

class GenerateDocumentationPages extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'docto:generate-pages';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Generates Documentation Pages with info on all the available commands.';

    /**
     * Execute the console command.
     */
    public function handle()
    {

    }
}
