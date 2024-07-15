<?php

use Illuminate\Support\Facades\Route;
use App\Http\Controllers\WordController;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider and all of them will
| be assigned to the "web" middleware group. Make something great!
|
*/

Route::get('/', function () {
    return view('welcome');
});

Route::get('word', [WordController::class, 'form_2']);
Route::get('word-2', [WordController::class, 'org']);
Route::get('word-3', [WordController::class, 'assessment']);
Route::get('word-4', [WordController::class, 'executive']);
Route::get('word-5', [WordController::class, 'appPath']);
Route::get('word-6', [WordController::class, 'appPathTable']);
