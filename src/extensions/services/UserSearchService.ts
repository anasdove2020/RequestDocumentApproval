import { IComboBoxOption } from "@fluentui/react";

export interface IUser {
  id: string;
  mail: string;
  displayName: string;
  department?: string;
  jobTitle?: string;
}

// Extended mock data for more realistic search results
// const MOCK_USERS: IUser[] = [
//   { id: "1", email: "john.doe@company.com", displayName: "John Doe", department: "Engineering", title: "Senior Developer" },
//   { id: "2", email: "jane.smith@company.com", displayName: "Jane Smith", department: "Marketing", title: "Marketing Manager" },
//   { id: "3", email: "mike.johnson@company.com", displayName: "Mike Johnson", department: "Engineering", title: "Tech Lead" },
//   { id: "4", email: "sarah.wilson@company.com", displayName: "Sarah Wilson", department: "HR", title: "HR Manager" },
//   { id: "5", email: "david.brown@company.com", displayName: "David Brown", department: "Finance", title: "Financial Analyst" },
//   { id: "6", email: "lisa.davis@company.com", displayName: "Lisa Davis", department: "Engineering", title: "Frontend Developer" },
//   { id: "7", email: "robert.taylor@company.com", displayName: "Robert Taylor", department: "Sales", title: "Sales Director" },
//   { id: "8", email: "emily.anderson@company.com", displayName: "Emily Anderson", department: "Design", title: "UX Designer" },
//   { id: "9", email: "james.martinez@company.com", displayName: "James Martinez", department: "Engineering", title: "DevOps Engineer" },
//   { id: "10", email: "maria.garcia@company.com", displayName: "Maria Garcia", department: "Product", title: "Product Manager" },
//   { id: "11", email: "william.lee@company.com", displayName: "William Lee", department: "Engineering", title: "Backend Developer" },
//   { id: "12", email: "jennifer.white@company.com", displayName: "Jennifer White", department: "Legal", title: "Legal Counsel" },
//   { id: "13", email: "michael.clark@company.com", displayName: "Michael Clark", department: "Operations", title: "Operations Manager" },
//   { id: "14", email: "amanda.rodriguez@company.com", displayName: "Amanda Rodriguez", department: "Marketing", title: "Content Specialist" },
//   { id: "15", email: "christopher.lewis@company.com", displayName: "Christopher Lewis", department: "Finance", title: "Controller" },
// ];

export class UserSearchService {
  /**
   * Simulates searching for users based on a query string
   * @param query - The search query (name, email, or department)
   * @param maxResults - Maximum number of results to return (default: 10)
   * @returns Promise that resolves to an array of matching users
   */
  public static async searchUsers(
    users: IUser[],
    query: string,
    maxResults: number = 10
  ): Promise<IUser[]> {
    // Simulate API delay
    await new Promise((resolve) =>
      setTimeout(resolve, 300 + Math.random() * 200)
    );

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

  /**
   * Converts user objects to ComboBox options
   * @param users - Array of user objects
   * @returns Array of ComboBox options
   */
  public static usersToComboBoxOptions(users: IUser[]): IComboBoxOption[] {
    return users.map((user) => ({
      key: user.mail,
      text: user.displayName,
      data: user, // Store full user data for later use
    }));
  }

  /**
   * Gets a user by email
   * @param mail - User's email address
   * @returns User object or undefined if not found
   */
  public static getUserByEmail(users: IUser[], mail: string): IUser | undefined {
    for (let i = 0; i < users.length; i++) {
      if (users[i].mail === mail) {
        return users[i];
      }
    }
    return undefined;
  }
}
