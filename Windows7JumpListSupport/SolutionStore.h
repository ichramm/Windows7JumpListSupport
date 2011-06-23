/*
 * Copyright (C) 2011 by Juan Ramirez (@ichramm) - jramirez@inconcertcc.com
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */
#ifndef solutionstore_h__
#define solutionstore_h__
#pragma once

#include <vector>

class CSolutionStore
{
public:

	CSolutionStore(void);

	~CSolutionStore(void);

	// get all solutions from the registry
	std::vector<CString> GetSolutions();

	// Saves a solution file in the registry
	int StoreSolution(const CString &path);

	// removes a solution file from the registry, maybe because it does not exist anymore
	int RemoveSolution(const CString &path);
};

#endif // solutionstore_h__
