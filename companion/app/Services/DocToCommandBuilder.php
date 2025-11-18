<?php

  namespace App\Services;

  class DocToCommandBuilder
  {
      public $params = [];

        public static function docto()
        {
            return (new static())->add(config('services.docto.path'));
        }

        public function add($param, $value = null){
            $this->params[] = [$param,$value];
            return $this;
        }

        public function build(){
            $cmd = '';
            foreach($this->params as $paramset){
                $cmd .= $paramset[0] . ' ' . $paramset[1] . ' ';
            }
            return $cmd;
        }
  }
