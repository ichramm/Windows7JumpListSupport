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
