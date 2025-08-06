export function getPriorityText(priority: number): string {
    switch (priority) {
        case 0: return 'No priority';
        case 1: return 'Urgent';
        case 3: return 'Important';
        case 5: return 'Medium';  
        case 9: return 'Low';
        default: return `Priority ${priority}`;
    }
}