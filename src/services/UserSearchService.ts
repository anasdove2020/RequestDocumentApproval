import { IComboBoxOption } from "@fluentui/react";
import { IUserSearch } from "../interfaces/IUserSearch";

export class UserSearchService {
  public static async searchUsers(
    users: IUserSearch[],
    query: string,
    maxResults: number = 10
  ): Promise<IUserSearch[]> {
    if (!query || query.length < 2) {
      return [];
    }

    const lowerQuery = query.toLowerCase();

    const filteredUsers = users.filter(
      (user) =>
        (user.displayName && user.displayName.toLowerCase().indexOf(lowerQuery) !== -1) ||
        (user.mail && user.mail.toLowerCase().indexOf(lowerQuery) !== -1) ||
        (user.department &&
          user.department.toLowerCase().indexOf(lowerQuery) !== -1) ||
        (user.jobTitle && user.jobTitle.toLowerCase().indexOf(lowerQuery) !== -1)
    );

    return filteredUsers.slice(0, maxResults);
  }
  
  public static usersToComboBoxOptions(users: IUserSearch[]): IComboBoxOption[] {
    return users.map((user) => ({
      key: user.mail,
      text: user.displayName,
      data: user,
    }));
  }

  public static getUserByEmail(users: IUserSearch[], mail: string): IUserSearch | undefined {
    for (let i = 0; i < users.length; i++) {
      if (users[i].mail === mail) {
        return users[i];
      }
    }
    return undefined;
  }
}
