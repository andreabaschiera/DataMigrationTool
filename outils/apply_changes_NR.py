import pandas as pd

def apply_changes_NR(df, version_cell, version_cell_new):

    dict_of_changes_NR = {
        "1.0.0": ["NumberOfPermanentContactEmployees", "DescriptionOfTheEffectiveParticipationOfWorkersUsersOrOtherinterestedPartiesOrCommunitiesInGovernance", "MostSeniorLevelAccountableForImplementationOfPracticesPoliciesAndOrFutureInitiatives"],
        "1.0.1": ["NumberOfPermanentContactEmployees", "DescriptionOfTheEffectiveParticipationOfWorkersUsersOrOtherinterestedPartiesOrCommunitiesInGovernance", "MostSeniorLevelAccountableForImplementationOfPracticesPoliciesAndOrFutureInitiatives"],
        "1.1.0": ["NumberOfPermanentContractEmployees", "DescriptionOfTheEffectiveParticipationOfWorkersUsersOrOtherInterestedPartiesOrCommunitiesInGovernance", "MostSeniorLevelAccountableForImplementationOfPolicies"],
        "1.1.1": ["NumberOfPermanentContractEmployees", "DescriptionOfTheEffectiveParticipationOfWorkersUsersOrOtherInterestedPartiesOrCommunitiesInGovernance", "MostSeniorLevelAccountableForImplementationOfPolicies"]
    }

    df_of_changes_NR = pd.DataFrame({
        "name_ranges": dict_of_changes_NR[version_cell],
        "name_ranges_new": dict_of_changes_NR[version_cell_new]
    })

    df = df.merge(df_of_changes_NR, on="name_ranges", how="left")

    for i in range(len(df)):
        if pd.notna(df["name_ranges_new"][i]):
            df["name_ranges"][i] = df["name_ranges_new"][i]
    
    df = df.drop(columns=["name_ranges_new"])

    return df