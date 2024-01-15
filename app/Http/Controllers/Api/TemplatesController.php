<?php

namespace App\Http\Controllers\Api;

use App\Http\Controllers\Controller;
use Illuminate\Http\Request;
use App\Utils\TemplateProcessor2;


class TemplatesController extends Controller
{
    /**
     * Display a listing of the resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function index()
    {
        $filepath = public_path('templates/template.docx');
        return response()->download($filepath, 'template.docx');
    }

    /**
     * Store a newly created resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @return \Illuminate\Http\Response
     */
    public function store(Request $request)
    {
        //
    }

    /**
     * Display the specified resource.
     *
     * @param  int  $id
     * @return \Illuminate\Http\Response
     */
    public function show($id)
    {
        //
    }

    /**
     * Update the specified resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @param  int  $id
     * @return \Illuminate\Http\Response
     */
    public function update(Request $request, $id)
    {
        //
    }

    /**
     * Remove the specified resource from storage.
     *
     * @param  int  $id
     * @return \Illuminate\Http\Response
     */
    public function destroy($id)
    {
        //
    }

    public function fill(Request $request)
    {
        $filepath = public_path('templates/template.docx');

        $phpword = new TemplateProcessor2($filepath);

        $phpword->replaceBookmark('full_name', 'אסי מון');
        $phpword->replaceBookmark('email', 'asim@example.com');
        $phpword->replaceBookmark('phone_no', '0583245657');
        $phpword->replaceBookmark('house_no', '12');

        $phpword->saveAs(public_path('templates/edited.docx'));
    }
}
